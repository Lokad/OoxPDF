using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    internal static void CheckComposedRunPageBreaks()
    {
        foreach (bool withNote in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = withNote ? "footnote-image-math" : "inline-math-navigation";
            var document = ReadMutation("CompositionProbes/" + fixture, mode, false, false, (xml, w, m) =>
            {
                XElement body = xml.Root!.Element(w + "body")!;
                XElement paragraph = body.Elements(w + "p").Single(p => p.Elements(m + "oMath").Count() == 2);
                XElement[] equations = paragraph.Elements(m + "oMath").ToArray();
                foreach (var pair in equations.Zip(new[] { "flowAlpha", "flowBeta" }))
                {
                    XElement run = pair.First.Element(m + "f")!.Element(m + "num")!.Element(m + "r")!;
                    run.Element(m + "t")!.Value = pair.Second;
                    if (run.Element(m + "rPr") is not { } properties)
                        run.AddFirst(new XElement(m + "rPr", new XElement(m + "nor")));
                    else if (properties.Element(m + "nor") is null) properties.Add(new XElement(m + "nor"));
                }
                equations[1].AddBeforeSelf(new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))));
                if (!withNote)
                {
                    // Retain the existing raster relation and page properties;
                    // the note case below keeps the complete composed document.
                    body.Elements().Where(child => !ReferenceEquals(child, paragraph) && child.Name != w + "sectPr").Remove();
                    xml.Descendants(w + "headerReference").Remove();
                    xml.Descendants(w + "footerReference").Remove();
                    return;
                }

                var children = paragraph.Elements().ToArray();
                int first = Array.IndexOf(children, equations[0]), last = Array.IndexOf(children, equations[1]);
                var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", "BreakCalculation"));
                equations[0].AddBeforeSelf(link);
                foreach (XElement child in children[first..(last + 1)]) { child.Remove(); link.Add(child); }
                link.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "900"), new XAttribute(w + "name", "BreakCalculation")));
                link.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "900")));
                body.Elements(w + "p").First().Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "BreakCalculation"),
                    new XElement(w + "r", new XElement(w + "t", "Jump to calculation"))));
            });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            string[] text = pages.Select(page => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: true).Select(show => show.Text))).ToArray();
            // A consumed wrapping space need not have its own painted glyph.
            // Check prose characters independently of line/show boundaries.
            string[] prose = text.Select(value => string.Concat(value.Where(c => !char.IsWhiteSpace(c)))).ToArray();
            TestAssert.Contains("flowAlpha", text[0]);
            TestAssert.True(!text[1].Contains("flowAlpha", StringComparison.Ordinal), "The first formula stays before the break.");
            TestAssert.Contains("flowBeta", text[1]);
            TestAssert.True(!text[0].Contains("flowBeta", StringComparison.Ordinal), "The second formula follows the break.");
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "A run break cannot discard supported formulas.");
            int Draws(PdfPage page) => page.Images.Sum(image => Regex.Matches(page.Content,
                @"/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count);
            TestAssert.Equal(1, Draws(pages[0]));
            TestAssert.Equal(0, Draws(pages[1]));
            if (!withNote) continue;
            TestAssert.Contains("Publiccalculationnote.", prose[0]);
            TestAssert.True(!prose[1].Contains("Publiccalculationnote.", StringComparison.Ordinal), "The note stays with its source marker.");
            TestAssert.True(prose[0].Contains("Endofsupportednote.", StringComparison.Ordinal), "The complete note tail remains visible: " + text[0]);
            TestAssert.Contains("Followingprose.", prose[1]);
            var target = pages[1].Annotations.Where(link => link.IsDestination).Single().Destination!.Value;
            TestAssert.Equal(0, target.PageIndex);
            TestAssert.True(pages[0].Annotations.Count(link => link.Destination == target) >= 2,
                "The leading navigation and first linked formula retain the bookmark destination.");
            TestAssert.True(pages[0].Annotations.Any(link => link.Destination is { PageIndex: 0 } note && note != target),
                "The note reference retains its independent destination.");
        }
    }
}
