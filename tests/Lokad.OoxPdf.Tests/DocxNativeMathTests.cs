using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNativeMathTests
{
    public static void DocxNativeDisplayMathRetainsMeaningAndReportsPaintedContent()
    {
        RequireMathFont();
        DocxDocument document = Fixture("literal", "after");
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains("Before formula:", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.True(page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "Native math must own an embedded font resource.");
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION" && diagnostic.PartName == "/word/document.xml"), "Painted math must be reported as approximate with source provenance.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"), "A fully painted source part must not report ignored mathematics.");
    }

    public static void DocxNativeMathUnsupportedStructuredRowsKeepTheirFallback()
    {
        RequireMathFont();
        foreach ((string kind, string placement) in new[] { ("scripts", "inline"), ("fraction", "inline") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(Fixture(kind, placement), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "Unsupported math must not publish unused equation fonts.");
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION" && diagnostic.Fallback == "Ignored"), "Unsupported structures retain an explicit ignored fallback.");
            TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_APPROXIMATED_EQUATION"), "Unpainted mathematics cannot claim native rendering.");
        }
    }

    public static void DocxNativeDisplayScriptsPaintDistinctBaseAndScriptRoles()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(Fixture("scripts", "after"), diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("xi2", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.Contains(" 8.52 Tf", page.Content);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "Painted source roles and script-style glyphs must report native approximate math.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "A fully painted display script cannot report ignored source content.");
    }

    public static void DocxNativeDisplayScriptsOwnTheirSelectedVariantSubsetMappings()
    {
        RequireMathFont();
        PdfPage page = Renderer().RenderBlankPages(Fixture("scripts", "after"), null, CancellationToken.None).Single();
        PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        foreach (int point in new[] { 0x1D456, (int)'2' })
        {
            ushort normal = resource.Font.Font.MapCodePoint(point);
            TestAssert.True(resource.Font.Font.TryGetMathScriptGlyph(normal, 1, CancellationToken.None, out ushort chosen),
                "Native script roles must resolve their source font's script-style glyph.");
            TestAssert.True(resource.Font.TryGetEncodedCid(chosen, out ushort cid), "The selected variant must own a published subset CID.");
            TestAssert.Equal(point, resource.Font.UnicodeByCid[cid]);
        }
    }

    public static void DocxNativeDisplayIndividualScriptsRetainTheirDistinctMeaning()
    {
        RequireMathFont();
        foreach ((string name, string expected) in new[] { ("docx-omml-script-sub12", "xi"), ("docx-omml-script-sup12", "x2") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture(name), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(expected, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
                "Single source-script roles must report their actual native painting.");
        }
    }

    public static void DocxNativeDisplayScriptFrameAdmitsItsFollowingProseAtTheBodyBoundary()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-script-boundary12");
        DocxRenderer renderer = Renderer();
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("xi2", text);
        TestAssert.Contains("Following prose.", text);
        var tail = renderer.InspectTextEmission(document).Lines.Last(line => line.SourceBlockIndex == 7);
        TestAssert.True(tail.Segments.All(segment => segment.BaselineY >= document.MarginBottomPoints),
            "The shared script frame must admit following prose without crossing the source body boundary.");
    }

    public static void DocxNativeDisplayFractionPaintsDistinctSourceRolesAndStoredRule()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(Fixture("fraction", "after"), diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("a+bc", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.Contains(" re f", page.Content);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "Painted numerator, denominator and rule must report approximate native rendering.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "A fully painted display fraction cannot report ignored source content.");
    }

    public static void DocxNativeDisplayFractionUsesItsFullFrameBeforeFollowingProseTurnsPage()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-display-fraction-boundary12");
        PdfPage[] pages = Renderer().RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.Equal(2, pages.Length);
        string first = DocxPageFootnoteNumberingTests.DecodeText(pages[0]).Normalize(NormalizationForm.FormKC);
        string second = DocxPageFootnoteNumberingTests.DecodeText(pages[1]).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("a+bc", first);
        TestAssert.True(!first.Contains("Following prose.", StringComparison.Ordinal),
            "Fraction ink, font leading and denominator must consume one complete frame before following prose is admitted.");
        TestAssert.Contains("Following prose.", second);
    }

    public static void DocxNativeInlineMathRetainsLiteralMeaningAndSharedProseBaseline()
    {
        RequireMathFont();
        DocxDocument document = Fixture("literal", "inline");
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains("Before formula:", text);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"), "Fitting inline content must report native approximate painting.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"), "Painted inline content cannot report ignored math.");
    }

    public static void DocxNativeInlineMathMovesFollowingProsePastItsMeasuredBox()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-inline-middle12");
        DocxRenderer renderer = Renderer();
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains(" after formula.", text);
        var segments = renderer.InspectTextEmission(document).Lines[0].Segments;
        double begin = segments[0].X;
        TestAssert.True(segments.Where(segment => !segment.IsTerminalLineSpace).Max(segment => segment.X) > begin + 120d,
            "Prose after a source-owned formula must move beyond its mathematical box.");
        TestAssert.True(segments.Where(segment => segment.IsTerminalLineSpace).All(segment => segment.X > begin + 200d),
            "The paragraph terminal cannot overlap inline mathematical content or following prose.");
    }

    public static void DocxNativeInlineMathOverfullRowsRetainTheWholeIgnoredFallback()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-inline-overfull12"), diagnostics.Add, CancellationToken.None).Single();
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"), "Formula-driven wrapping retains an explicit fallback until qualified.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_APPROXIMATED_EQUATION"), "An overfull row cannot claim painted math.");
        TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "Fallback output cannot publish unused math fonts.");
    }

    public static void DocxNativeInlineMathWrapsFollowingProseWithoutLosingItsSourceBoundary()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-wrap-middle140");
        var diagnostics = new List<OoxPdfDiagnostic>();
        DocxRenderer renderer = Renderer();
        PdfPage page = renderer.RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains("after formula.", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.True(!text.Contains('\u200B'), "Layout markers cannot leak into emitted prose.");
        var lines = renderer.InspectTextEmission(document).Lines.Where(line => line.SourceBlockIndex == 0).ToArray();
        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines[0].Segments[0].BaselineY > lines[1].Segments[0].BaselineY,
            "Following prose must own a subsequent line, preserving the mathematical source position.");
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "A fitting formula remains painted when only its following prose requires wrapping.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "Supported prose wrapping cannot claim ignored mathematical content.");
    }

    public static void DocxNativeInlineMathWrapsBeginningAndLongTailProseWithoutChangingSourceText()
    {
        RequireMathFont();
        foreach (string name in new[] { "docx-omml-prosewrap-begin90", "docx-omml-prosewrap-long140" })
        {
            DocxDocument document = NamedFixture(name);
            DocxRenderer renderer = Renderer();
            PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
            string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
            TestAssert.Contains("α+β=γ", text);
            TestAssert.Contains("Following prose.", text);
            string prose = string.Concat(document.Paragraphs[0].Runs.Select(run => run.Text));
            TestAssert.True(prose.Split(' ', StringSplitOptions.RemoveEmptyEntries).All(word => text.Contains(word, StringComparison.Ordinal)),
                "Every following prose word must survive source-owned formula wrapping.");
            var lines = renderer.InspectTextEmission(document).Lines.Where(line => line.SourceBlockIndex == 0).ToArray();
            TestAssert.True(lines.Length >= 2, "The mathematical box must participate in following-prose line breaking.");
        }
    }

    public static void DocxNativeMathSettingsAndExplicitRunFacesRemainDistinct()
    {
        DocxDocument document = Fixture("literal", "after");
        TestAssert.Equal("Cambria Math", document.Settings.MathFontFamily);
        DocxEquation equation = document.Paragraphs.Single(paragraph => paragraph.Equations.Count != 0).Equations.Single();
        TestAssert.True(equation.Content.Children.Cast<DocxMathRun>().All(run => run.ExplicitFontFamily == "Cambria Math"), "Authored math faces must remain separate from inherited prose faces.");
    }

    public static void DocxNativeMathUsesItsOccupiedBodyFrameBeforeFollowingProseTurnsPage()
    {
        RequireMathFont();
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-omml-display-boundary12.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        PdfPage[] pages = Renderer().RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.Equal(2, pages.Length);
        string first = DocxPageFootnoteNumberingTests.DecodeText(pages[0]).Normalize(NormalizationForm.FormKC);
        string second = DocxPageFootnoteNumberingTests.DecodeText(pages[1]).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", first);
        TestAssert.True(!first.Contains("Following prose.", StringComparison.Ordinal), "The following paragraph must turn at the source body boundary.");
        TestAssert.Contains("Following prose.", second);
    }

    public static void DocxNativeMathMissingFontsAndUnqualifiedFormattingRetainIgnoredDiagnostics()
    {
        DocxDocument document = Fixture("literal", "after");
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = new DocxRenderer(new MapFontResolver([], "Unavailable"), OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"), "Unavailable math fonts retain ignored-source diagnostics.");
        TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "A missing math face cannot claim embedded equation output.");
        DocxParagraph paragraph = document.Paragraphs.Single(value => value.Equations.Count != 0);
        TestAssert.NotNull(DocxMathParagraphPolicy.GetDisplayRow(paragraph));
        TestAssert.True(DocxMathParagraphPolicy.GetDisplayRow(paragraph with
        { Spacing = paragraph.Spacing with { LineValue = "480", LineRuleValue = "exact" } }) is null,
            "Unqualified explicit math line rules must keep their existing fallback.");
        TestAssert.True(DocxMathParagraphPolicy.GetDisplayRow(paragraph with
        { InlineReferences = [new DocxInlineReference(DocxRelatedStoryKind.Footnote, "1", null, "1", 0, 0, 0)] }) is null,
            "References require their own semantic owner before math painting is admitted.");
    }

    private static DocxRenderer Renderer() => new(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
    private static DocxDocument Fixture(string kind, string placement) => DocxTests.ReadDocx(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", $"docx-omml-{kind}-{placement}.docx")), OoxPdfDocxMarkupMode.Final);
    private static DocxDocument NamedFixture(string name) => DocxTests.ReadDocx(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name + ".docx")), OoxPdfDocxMarkupMode.Final);
    private static void RequireMathFont()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for this native rendering control.");
    }
}
