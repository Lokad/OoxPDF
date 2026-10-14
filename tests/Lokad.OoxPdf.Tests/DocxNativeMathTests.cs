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

    public static void DocxNativeMathUnsupportedInlineAndStructuredRowsKeepTheirFallback()
    {
        RequireMathFont();
        foreach ((string kind, string placement) in new[] { ("literal", "inline"), ("scripts", "after"), ("fraction", "after") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(Fixture(kind, placement), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "Unsupported math must not publish unused equation fonts.");
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION" && diagnostic.Fallback == "Ignored"), "Unsupported structures retain an explicit ignored fallback.");
            TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_APPROXIMATED_EQUATION"), "Unpainted mathematics cannot claim native rendering.");
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
    private static void RequireMathFont()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for this native rendering control.");
    }
}
