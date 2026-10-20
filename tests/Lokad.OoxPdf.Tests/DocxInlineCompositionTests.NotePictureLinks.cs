using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNotePictureLinksKeepClippedMissingAndActionFallbacks()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        foreach (string variant in new[] { "clipped", "missing", "action" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("Composition/composed-picture-navigation-report", mode, false, false,
                (xml, w, _) =>
                {
                    if (variant == "missing") xml.Descendants(a + "blip").Single().SetAttributeValue(r + "embed", "missing");
                    if (variant == "action") xml.Descendants(a + "hlinkClick").Single().SetAttributeValue("action", "unsupported");
                    if (variant == "clipped")
                    {
                        var paragraph = xml.Descendants(w + "p").Single(p => p.Descendants(a + "blip").Any());
                        var properties = paragraph.Element(w + "pPr");
                        if (properties is null) { properties = new XElement(w + "pPr"); paragraph.AddFirst(properties); }
                        var spacing = properties.Element(w + "spacing");
                        if (spacing is null) { spacing = new XElement(w + "spacing"); properties.Add(spacing); }
                        spacing.SetAttributeValue(w + "line", "20");
                        spacing.SetAttributeValue(w + "lineRule", "exact");
                    }
                }, "word/endnotes.xml");
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
            var links = pages.SelectMany(page => page.Annotations).Where(link =>
                link.Uri?.StartsWith("https://example.invalid/composed-picture/", StringComparison.Ordinal) == true).ToArray();
            TestAssert.Equal(0, links.Count(link => link.Uri!.Contains("/endnotes/", StringComparison.Ordinal)));
            TestAssert.Equal(2 + 2 * pages.Length, links.Length);
            TestAssert.Equal(2 + 2 * pages.Length + (variant == "missing" ? 0 : 1), pages.Sum(page => page.Images.Count));
            TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Picture navigation fallbacks preserve neighboring supported equations.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(source, null, cancellation.Token).ToArray());
        }
    }
}
