using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDisabledStaticFontsStayInspectableWithoutResolverCalls()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "first", "even" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadStaticResourceFontSource(story, variant, mode, enabled: false);
            TestAssert.True(DocxBlockTraversal.EnumerateSourceParagraphs(source).Any(p => p.Runs.Any(r => r.FontFamily == "Consolas")),
                "All-source inspection retains the deliberately unprinted font owner.");
            var resolver = new StaticFontRequestResolver("Consolas", "Cambria Math");
            var page = new DocxRenderer(resolver, mode, CachedMathGeometry(mode)).RenderPagesAsync(source, null, default)
                .ToBlockingEnumerable().Single();
            TestAssert.Contains("Before public math table.", DocxPageFootnoteNumberingTests.DecodeText(page));
            TestAssert.Contains("After public math table.", DocxPageFootnoteNumberingTests.DecodeText(page));
            TestAssert.True(resolver.Requests.All(name => name is not ("Consolas" or "Cambria Math")),
                "Neither text nor math preparation can request an unreachable static font.");
        }
    }

    public static void DocxActiveStaticTextAndMathFontsRetainTerminalFailures()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "first", "even" })
        foreach (string refused in new[] { "Consolas", "Cambria Math" })
        {
            var source = ReadStaticResourceFontSource(story, variant, OoxPdfDocxMarkupMode.Final, enabled: true);
            if (variant == "even") source = source with { BodyElements = source.BodyElements.Concat<DocxBodyElement>(
                [new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Second page body.", 12d, 12d))]).ToArray() };
            var resolver = new StaticFontRequestResolver(refused);
            TestAssert.Throws<InvalidOperationException>(() => new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray());
            TestAssert.True(resolver.Requests.Contains(refused), "A required text/math font refusal must still propagate.");
        }
    }

    public static void DocxInheritedStaticFontsKeepTheirActualFaceAndFailures()
    {
        foreach (string story in new[] { "header", "footer" })
        {
            var source = ReadInheritedFirstStory(story, OoxPdfDocxMarkupMode.Final, formattedTable: false);
            var strict = new StaticFontRequestResolver("Consolas");
            TestAssert.Throws<InvalidOperationException>(() => new DocxRenderer(strict, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray());
            TestAssert.True(strict.Requests.Contains("Consolas"), "A later enabled inherited story still requires its font.");
            var pages = new DocxRenderer(new SectionBoxResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            string label = "SECTION " + story.ToUpperInvariant() + " BOX LABEL";
            string actualFaceText = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[1], true)
                .Where(show => show.Font.Font.FamilyName == "DejaVu Math TeX Gyre").Select(show => show.Text));
            TestAssert.Contains(label, actualFaceText);
        }
    }

    public static void DocxDisabledStaticSvgFontsUseSharedResourceReachability()
    {
        var source = ReadStaticVariantPropertySource("header", "first", OoxPdfDocxMarkupMode.Final, history: false, enabled: false);
        var paragraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 12d, 12d);
        var image = new DocxInlineImage(60d, 20d, "image/svg+xml", Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"60\" height=\"20\"><text x=\"1\" y=\"15\" font-family=\"Consolas\">LABEL</text></svg>"),
            "/word/media/static-label.svg") { SourceRunIndex = 0 };
        paragraph = paragraph with { Images = [image], InlineContent = [new DocxInlinePicture(image)] };
        var settings = source.PageSettings with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["first"] = DocxTests.CreateParagraphBlocks(paragraph) }
        };
        source = WithStaticVariantSettings(source, settings) with { EquationParts = [], EquationDiagnostics = [] };
        var resolver = new StaticFontRequestResolver("Consolas");
        var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(source, null, default)
            .ToBlockingEnumerable().Single();
        TestAssert.Contains("Before public math table.", DocxPageFootnoteNumberingTests.DecodeText(page));
        TestAssert.True(!resolver.Requests.Contains("Consolas"), "Unused SVG labels cannot acquire their font after pagination.");
        var active = WithStaticVariantSettings(source, settings with { TitlePage = true });
        var refused = new StaticFontRequestResolver("Consolas");
        TestAssert.Throws<InvalidOperationException>(() => new DocxRenderer(refused, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderPagesAsync(active, null, default).ToBlockingEnumerable().ToArray());
        TestAssert.True(refused.Requests.Contains("Consolas"), "The same SVG font remains required when its story is enabled.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxFontPlan.Create(source, resolver, cancellation.Token));
    }

    private static DocxDocument ReadStaticResourceFontSource(string story, string variant, OoxPdfDocxMarkupMode mode, bool enabled)
    {
        var source = ReadMutation("InlineBoxMath/inline-box-math-" + story, mode, false, false, (xml, w, _) =>
        {
            var run = xml.Descendants(w + "txbxContent").Single().Descendants(w + "r").First(r => r.Element(w + "t") is not null);
            var properties = run.Element(w + "rPr");
            if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
            var fonts = properties.Element(w + "rFonts");
            if (fonts is null) { fonts = new XElement(w + "rFonts"); properties.AddFirst(fonts); }
            foreach (string name in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) fonts.SetAttributeValue(w + name, "Consolas");
            foreach (string name in new[] { "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme" }) fonts.SetAttributeValue(w + name, null);
        }, partOverride: "word/inline-box-" + story + ".xml");
        return WithStaticVariantSettings(source, StaticVariantSettings(source.PageSettings, story, variant, enabled));
    }

    private sealed class StaticFontRequestResolver(params string[] refused) : IFontResolver
    {
        private readonly PortableRenderingTests.PinnedResolver inner = new(mathFont: true);
        public List<string> Requests { get; } = [];
        public FontFaceResolution Resolve(FontRequest request)
        {
            Requests.Add(request.FamilyName);
            if (refused.Contains(request.FamilyName, StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException("Expected required font refusal: " + request.FamilyName);
            return inner.Resolve(request);
        }
    }
}
