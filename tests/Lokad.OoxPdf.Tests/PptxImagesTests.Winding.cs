using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxImagesTests
{
    public static void SvgFillRulesFollowTheSourcePathAcrossPaintAndDocumentKinds()
    {
        foreach (bool pptx in new[] { true, false })
        foreach (string paint in new[] { "solid", "linearGradient", "radialGradient" })
        foreach (bool evenOdd in new[] { true, false })
        {
            string definition = paint == "solid" ? "" : $"<defs><{paint} id=\"g\"><stop offset=\"0\" stop-color=\"#DC2814\"/><stop offset=\"1\" stop-color=\"#DC2814\"/></{paint}></defs>";
            string fill = paint == "solid" ? "#DC2814" : "url(#g)";
            string svg = $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\">{definition}<path fill=\"{fill}\" fill-rule=\"{(evenOdd ? "evenodd" : "nonzero")}\" d=\"M0 0H100V100H0Z M25 25H75V75H25Z\"/></svg>";
            var entries = pptx ? JpegPassthroughTests.PptxEntries() : JpegPassthroughTests.DocxEntries();
            foreach (string name in entries.Keys.ToArray())
                entries[name] = TestFixtures.Utf8(Encoding.UTF8.GetString(entries[name])
                    .Replace("probe.jpg", "probe.svg", StringComparison.Ordinal)
                    .Replace("Extension=\"jpg\" ContentType=\"image/jpeg\"", "Extension=\"svg\" ContentType=\"image/svg+xml\"", StringComparison.Ordinal));
            entries[(pptx ? "ppt" : "word") + "/media/probe.svg"] = TestFixtures.Utf8(svg);
            using var input = TestFixtures.CreateZipPackage(entries);
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var resolver = new PortableRenderingTests.PinnedResolver();
            PdfPage page = pptx
                ? new PptxRenderer(resolver).RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default).ToBlockingEnumerable().Single()
                : new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderPagesAsync(new DocxReader().Read(package, diagnostics.Add, default, OoxPdfDocxMarkupMode.Final), diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(0, diagnostics.Count);
            // Bind to the closed source contour, excluding unrelated rectangular
            // page backgrounds and viewport clips. Native controls inspect its hole.
            MatchCollection operations = Regex.Matches(page.Content, @"(?m)^h\r?\n(?<paint>W\*? n|f\*?)\r?$");
            TestAssert.Equal(1, operations.Count);
            string expected = paint == "solid" ? evenOdd ? "f*" : "f" : evenOdd ? "W* n" : "W n";
            TestAssert.Equal(expected, operations[0].Groups["paint"].Value);
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }
}
