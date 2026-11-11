using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxPictureBookmarksWithoutMathRetainSourcePositions()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool wrapped in new[] { false, true })
        {
            var source = ReadPictureBookmark(container, mode, 18, wrapped, includeMath: false);
            var target = DocxBlockTraversal.EnumerateSourceParagraphs(source)
                .First(paragraph => paragraph.BookmarkAnchors.Any(anchor => anchor.Name == "PictureTarget"));
            var items = TestAssert.NotNull(target.InlineContent);
            var anchor = target.BookmarkAnchors.Single(anchor => anchor.Name == "PictureTarget");
            TestAssert.True(target.Equations.Count == 0 && target.Hyperlinks.Count == 0 &&
                anchor.InlineItemIndex is int index && index >= 0 && index < items.Count &&
                items[index] is DocxInlinePicture picture && ReferenceEquals(picture.Image, target.Images.Single()),
                "Picture source ownership does not depend on math or a link in the target paragraph.");
        }
    }

    public static void DocxPictureBookmarksWithoutMathUseTheSamePlacedSource()
    {
        // Ordinary placement and ordered static stories publish different
        // records, but both destinations must reveal their actual picture.
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool wrapped in new[] { false, true })
        {
            var source = ReadPictureBookmark(container, mode, 18, wrapped, includeMath: false);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var control = renderer.RenderPagesAsync(ReadPictureBookmark(container, mode, 18, wrapped,
                includeMath: false, linkPrefix: true), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(control.Length, pages.Length);
            var destination = pages.SelectMany(page => page.Annotations).Single(a => a.Destination is not null).Destination!.Value;
            var painted = pages.Select((page, index) => (Page: page, Index: index)).Single(pair => pair.Page.Images.Count != 0);
            var frame = ReadPaintedPictureFrame(painted.Page);
            double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source, DocxMarkupContext.FromMode(mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
            TestAssert.Equal(painted.Index, destination.PageIndex);
            TestAssert.True(Math.Abs(frame.X - destination.Left!.Value - 3d * scale) < .001d &&
                destination.Top!.Value >= frame.Top - .001d,
                $"The target uses its picture's X and occupied row top: {container}/{mode}/{wrapped}, scale {scale}, picture {frame}, destination {destination}.");
            for (int page = 0; page < pages.Length; page++)
                TestAssert.Equal(control[page].Content, pages[page].Content);
            foreach (var link in control.SelectMany(page => page.Annotations))
                TestAssert.Equal(destination, link.Destination!.Value);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (string value in new[] { "PublicPrefix", "PublicSuffix", "Jump to picture" }) TestAssert.Contains(value, text);
            TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
        }
    }

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
            var picture = ReadPaintedPictureFrame(painted.Page);
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
        foreach (bool includeMath in new[] { false, true })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = renderer.RenderPagesAsync(ReadPictureBookmark(container, mode, 18, false,
                includeMath: includeMath), null, default)
                .ToBlockingEnumerable().ToArray();
            var expected = control.SelectMany(page => page.Annotations).Single(a => a.Destination is not null).Destination!.Value;
            string expectedText = string.Concat(control.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (bool hidden in new[] { false, true })
            {
                var pages = renderer.RenderPagesAsync(ReadPictureBookmark(container, mode, 18, false, gap: true, hidden,
                    includeMath: includeMath), null, default)
                    .ToBlockingEnumerable().ToArray();
                var actual = pages.SelectMany(page => page.Annotations).Single(a => a.Destination is not null).Destination!.Value;
                TestAssert.Equal(control.Length, pages.Length);
                TestAssert.Equal(expectedText, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)));
                TestAssert.Equal(expected, actual);
                TestAssert.Equal(ReadPaintedPictureFrame(control.Single(p => p.Images.Count != 0)),
                    ReadPaintedPictureFrame(pages.Single(p => p.Images.Count != 0)));
            }
        }
    }

    internal static (double X, double Top) ReadPaintedPictureFrame(PdfPage page)
    {
        var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" +
            Regex.Escape(page.Images.Single().ResourceName) + @" Do");
        TestAssert.True(draw.Success, "The destination is compared with an actual painted picture.");
        double Number(int index) => double.Parse(draw.Groups[index].Value, CultureInfo.InvariantCulture);
        return (Number(3), Number(4) + Number(2));
    }

    private static DocxDocument ReadPictureBookmark(string container, OoxPdfDocxMarkupMode mode, int height,
        bool wrapped, bool gap = false, bool hidden = false, bool includeMath = true, bool linkPrefix = false) => ReadMutation(
        "CompositionProbes/" + (container == "cell" ? "body" : container), mode, false, false,
        (xml, w, m) =>
        {
            var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(m + "oMath").Any());
            var picture = paragraph.Elements(w + "r").Single(run => run.Descendants(w + "drawing").Any());
            foreach (var extent in picture.Descendants().Where(e => e.Name.LocalName is "extent" or "ext" && e.Attribute("cy") is not null))
                extent.SetAttributeValue("cy", height * 12700);
            var first = paragraph.Elements().First(element => element.Name != w + "pPr");
            var prefix = new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                "PublicPrefix " + (wrapped ? string.Concat(Enumerable.Repeat("Visible wrapping prefix ", 12)) : string.Empty)));
            first.AddBeforeSelf(linkPrefix ? new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PictureTarget"), prefix) : prefix);
            if (!includeMath)
                foreach (var equation in paragraph.Elements(m + "oMath").ToArray()) equation.Remove();
            picture.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "909"), new XAttribute(w + "name", "PictureTarget")));
            if (gap) picture.AddBeforeSelf(hidden
                ? new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "t", "HiddenProbe"))
                : new XElement(w + "r"));
            picture.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "909")));
            paragraph.Add(new XElement(w + "r", new XElement(w + "t", " PublicSuffix")));
            // Navigation from another paragraph requires placed source ranges
            // even when the optional prefix link is absent.
            paragraph.AddBeforeSelf(new XElement(w + "p", new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PictureTarget"),
                new XElement(w + "r", new XElement(w + "t", "Jump to picture")))));
            if (container == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
        }, partOverride: container == "header" ? "word/header1.xml" : container == "footer" ? "word/footer1.xml" : null);
}
