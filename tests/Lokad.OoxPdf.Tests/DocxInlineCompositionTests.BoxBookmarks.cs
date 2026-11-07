using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineBoxBookmarksPreservePlacedLinksAndSiblingContent()
    {
        foreach (string variant in new[] { "body", "first", "wrap", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = variant == "body" ? "inline-box-body-equation-inline-box-valid" : "inline-sibling-" + variant;
            string part = variant == "header" ? "word/inline-box-header.xml" :
                variant == "footer" ? "word/inline-box-footer.xml" : "word/document.xml";
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = renderer.RenderPagesAsync(ReadMutation("InlineBoxMath/" + fixture, mode, false, false,
                partOverride: part), null, default).ToBlockingEnumerable().ToArray();
            double? targetX = null, targetY = null;
            foreach (string gap in new[] { "none", "empty", "hidden" })
            {
                var source = ReadMutation("InlineBoxMath/" + fixture, mode, false, false, (xml, w, m) =>
                {
                    var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(m + "oMath").Any());
                    var box = paragraph.Elements(w + "r").Single(run => run.Descendants(w + "txbxContent").Any());
                    box.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "910"), new XAttribute(w + "name", "InlineBoxTarget")));
                    if (gap != "none") box.AddBeforeSelf(gap == "empty" ? new XElement(w + "r") :
                        new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "t", "HiddenProbe")));
                    box.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "910")));
                    var text = paragraph.Elements(w + "r").First(run => run.Element(w + "t")?.Value.Length > 0);
                    text.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "InlineBoxTarget"), new XElement(text)));
                }, part);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(control.Length, pages.Length);
                TestAssert.Equal(string.Concat(control.Select(DocxPageFootnoteNumberingTests.DecodeText)),
                    string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)));
                TestAssert.Equal(control.Sum(page => page.Images.Count), pages.Sum(page => page.Images.Count));
                var destination = pages.SelectMany(page => page.Annotations).Single(a => a.Destination is not null).Destination!.Value;
                TestAssert.True(destination.PageIndex >= 0 && destination.PageIndex < pages.Length &&
                    destination.Left is not null && destination.Top is > 0d,
                    "A supported inline box retains a concrete internal-link destination: " + variant + "/" + mode + "/" + gap);
                double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source, DocxMarkupContext.FromMode(mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                var page = pages[destination.PageIndex];
                var frames = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re\s+S");
                TestAssert.True(frames.Any(frame =>
                {
                    double Number(int index) => double.Parse(frame.Groups[index].Value, CultureInfo.InvariantCulture);
                    return Math.Abs(destination.Left.Value + 3d * scale - Number(1)) < .001d &&
                        Math.Abs(destination.Top!.Value - Number(2) - Number(4)) < .001d;
                }), "Navigation uses an actual painted box frame rather than surrounding prose.");
                if (targetX is not null)
                    TestAssert.True(Math.Abs(destination.Left.Value - targetX.Value) < .000001d &&
                        Math.Abs(destination.Top!.Value - targetY!.Value) < .000001d,
                        "Empty/hidden source runs cannot redirect the box target.");
                targetX = destination.Left; targetY = destination.Top;
            }
        }
    }
}
