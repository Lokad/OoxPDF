using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

// Real rendering with one pinned, redistributable face. These are content and
// resource checks on both platforms; Office visual comparisons remain separate.
internal static class PortableRenderingTests
{
    public static void PortableRenderingDocxFinalPreservesMixedContent() =>
        CheckDocx(OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

    public static void PortableRenderingDocxAllMarkupPreservesMixedContent() =>
        CheckDocx(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    public static void PortableRenderingPptxPreservesCompositeChartAndText()
    {
        var package = OoxPackage.Open(Case("pptx-ladder-11-composite-chart-port.pptx"), default);
        var document = new PptxReader().Read(package, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = new PptxRenderer(new PinnedResolver()).RenderPages(document, package, diagnostics.Add, default).ToArray();
        TestAssert.Equal(1, pages.Length);
        string text = DocxPageFootnoteNumberingTests.DecodeText(pages.Single());
        TestAssert.Contains("Sales Overview", text);
        TestAssert.Contains("Category 1", text);
        TestAssert.Contains("Category 4", text);
        TestAssert.Contains("Series 3", text);
        TestAssert.True(pages[0].Content.Contains(" rg", StringComparison.Ordinal), "The composite chart must also paint graphics.");
        CheckWrittenPages(pages);
    }

    private static void CheckDocx(OoxPdfDocxMarkupMode mode, OoxPdfDocxMarkupGeometryMode geometry)
    {
        var document = DocxTests.ReadDocx(Case("PortableFonts/letter-normal-math.docx"), mode);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = new DocxRenderer(new PinnedResolver(), mode, geometry)
            .RenderBlankPages(document, diagnostics.Add, default).ToArray();
        string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("Engineering review letter", text);
        TestAssert.Contains("Acceptance of the revised delivery schedule", text);
        TestAssert.Contains("All quantities in this synthetic example are illustrative.", text);
        TestAssert.Equal(3, text.Split("x+2y", StringSplitOptions.None).Length - 1);
        TestAssert.True(text.Split("a+b=c", StringSplitOptions.None).Length - 1 >= 3,
            "Body, cell and header equations must reach painting with portable fonts.");
        TestAssert.Equal(1, pages.Sum(page => page.Images.Sum(image => Regex.Matches(page.Content,
            "/" + Regex.Escape(image.ResourceName) + @"\s+Do(?:\s|$)").Count)));
        TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Supported equations must not disappear.");
        CheckWrittenPages(pages);
    }

    private static void CheckWrittenPages(PdfPage[] pages)
    {
        TestAssert.True(pages.Length > 0 && pages.All(page => page.Fonts.Count > 0), "Real embedded text resources are required.");
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, pages, default);
        string pdf = Encoding.Latin1.GetString(output.ToArray());
        TestAssert.Contains("/FontFile2", pdf);
        TestAssert.Contains("/ToUnicode", pdf);
        TestAssert.Contains("%%EOF", pdf);
    }

    private static string Case(string relative) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", relative));

    private sealed class PinnedResolver : IFontResolver
    {
        private readonly FileFontProgramSource source;

        public PinnedResolver()
        {
            string path = Case("PortableFonts/DejaVuSans.ttf");
            TestAssert.Equal("AE7B7855E115A5966D8B1B3F80F254CCC117EC86F9965E202EE2940453837280",
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            source = new(path);
        }

        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "DejaVu Sans",
            new FontStyleKey(HasMathTable: true), source, IsFallback: false);
    }
}
