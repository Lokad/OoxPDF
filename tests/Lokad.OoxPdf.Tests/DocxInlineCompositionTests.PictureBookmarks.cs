using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlinePictureBookmarksFollowPaintedFramesAcrossContainers()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int height in new[] { 6, 18, 36 })
        foreach (bool wrapped in new[] { false, true })
        {
            var source = ReadPictureBookmark(container, mode, height, wrapped);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var painted = pages.Select((page, index) => (Page: page, Index: index)).Single(pair => pair.Page.Images.Count != 0);
            var destination = pages.SelectMany(page => page.Annotations).Single(a => a.Destination is not null).Destination!.Value;
            var picture = ReadPictureBookmarkFrame(painted.Page);
            double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source, DocxMarkupContext.FromMode(mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
            TestAssert.Equal(painted.Index, destination.PageIndex);
            TestAssert.True(Math.Abs(picture.X - destination.Left!.Value - 3d * scale) < .001d,
                "The inbound link targets its actual picture, including wrapping and container translations: " + container + "/" + mode);
            TestAssert.True(destination.Top!.Value >= picture.Top - .001d,
                "A short picture's bookmark retains the occupied line top rather than its lower image top.");
            if (height >= 18)
                TestAssert.True(Math.Abs(destination.Top.Value - picture.Top) < .001d,
                    "A line dominated by its picture targets that painted top.");
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string value in new[] { "PublicPrefix", "PublicSuffix", "Jump to picture" }) TestAssert.Contains(value, text);
            TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
        }
    }

    public static void DocxInlinePictureBookmarksKeepTargetsAcrossNonpaintingRuns()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = renderer.RenderPagesAsync(ReadPictureBookmark(container, mode, 18, false), null, default)
                .ToBlockingEnumerable().ToArray();
            var expected = control.SelectMany(page => page.Annotations).Single(a => a.Destination is not null).Destination!.Value;
            string expectedText = string.Concat(control.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (bool hidden in new[] { false, true })
            {
                var pages = renderer.RenderPagesAsync(ReadPictureBookmark(container, mode, 18, false, gap: true, hidden), null, default)
                    .ToBlockingEnumerable().ToArray();
                var actual = pages.SelectMany(page => page.Annotations).Single(a => a.Destination is not null).Destination!.Value;
                TestAssert.Equal(control.Length, pages.Length);
                TestAssert.Equal(expectedText, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)));
                TestAssert.Equal(expected, actual);
                TestAssert.Equal(ReadPictureBookmarkFrame(control.Single(p => p.Images.Count != 0)),
                    ReadPictureBookmarkFrame(pages.Single(p => p.Images.Count != 0)));
            }
        }
    }

    private static (double X, double Top) ReadPictureBookmarkFrame(PdfPage page)
    {
        var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" +
            Regex.Escape(page.Images.Single().ResourceName) + @" Do");
        TestAssert.True(draw.Success, "The destination is compared with an actual painted picture.");
        double Number(int index) => double.Parse(draw.Groups[index].Value, CultureInfo.InvariantCulture);
        return (Number(3), Number(4) + Number(2));
    }

    private static DocxDocument ReadPictureBookmark(string container, OoxPdfDocxMarkupMode mode, int height,
        bool wrapped, bool gap = false, bool hidden = false) => ReadMutation(
        "CompositionProbes/" + (container == "cell" ? "body" : container), mode, false, false,
        (xml, w, m) =>
        {
            var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(m + "oMath").Any());
            var picture = paragraph.Elements(w + "r").Single(run => run.Descendants(w + "drawing").Any());
            foreach (var extent in picture.Descendants().Where(e => e.Name.LocalName is "extent" or "ext" && e.Attribute("cy") is not null))
                extent.SetAttributeValue("cy", height * 12700);
            var first = paragraph.Elements().First(element => element.Name != w + "pPr");
            first.AddBeforeSelf(new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                "PublicPrefix " + (wrapped ? string.Concat(Enumerable.Repeat("Visible wrapping prefix ", 12)) : string.Empty))));
            picture.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "909"), new XAttribute(w + "name", "PictureTarget")));
            if (gap) picture.AddBeforeSelf(hidden
                ? new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "t", "HiddenProbe"))
                : new XElement(w + "r"));
            picture.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "909")));
            paragraph.Add(new XElement(w + "r", new XElement(w + "t", " PublicSuffix")));
            // The target paragraph deliberately has no hyperlinks or comments;
            // navigation from a different paragraph still requires placed source ranges.
            paragraph.AddBeforeSelf(new XElement(w + "p", new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PictureTarget"),
                new XElement(w + "r", new XElement(w + "t", "Jump to picture")))));
            if (container == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
        }, partOverride: container == "header" ? "word/header1.xml" : container == "footer" ? "word/footer1.xml" : null);
}
