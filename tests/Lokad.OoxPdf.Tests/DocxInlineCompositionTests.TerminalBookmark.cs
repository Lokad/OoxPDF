using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxTerminalBookmarksSurviveUnpaintedWhitespaceBeforePageBreak()
    {
        foreach (string fixture in new[] { "CompositionProbes/math-inner-bookmark-outer", "Composition/composed-static-cell-report" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            string name = fixture.StartsWith("Composition/", StringComparison.Ordinal) ? "ComposedCalculation" : "MathSource";
            var document = ReadMutation(fixture, mode, false, false, (xml, w, _) =>
            {
                var start = xml.Descendants(w + "bookmarkStart").Single(value => (string?)value.Attribute(w + "name") == name);
                start.AddAfterSelf(new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))));
            });
            var owner = document.Paragraphs.Single(paragraph => paragraph.BookmarkAnchors.Any(value => value.Name == name));
            TestAssert.True(owner.Equations.Count == 0 && owner.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text)),
                "The source bookmark remains at the text fragment before the explicit break.");
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(pages.Length >= 2 && pages.SelectMany(page => page.Annotations)
                .Any(link => link.Destination is { PageIndex: 0, Left: not null, Top: not null }),
                "The internal link must retain a concrete destination at the preceding painted source endpoint: " + fixture + "/" + mode);
        }
    }
}
