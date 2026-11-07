using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFirstMathBookmarkKeepsItsBoundaryAcrossNonpaintingRuns()
    {
        foreach (string context in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var variation in new[] { (Gap: 0, Hidden: false), (Gap: 1, Hidden: false), (Gap: 3, Hidden: false),
            (Gap: 1, Hidden: true), (Gap: 3, Hidden: true) })
        {
            var paragraph = NonpaintingBookmarkParagraph(ReadNonpaintingBookmark(context, variation.Gap, mode, hidden: variation.Hidden));
            var anchor = paragraph.BookmarkAnchors.Single(a => a.Name == "PrefixMathTarget");
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.True(anchor.InlineItemIndex is int index && index >= 0 && index < items.Count &&
                items[index] is DocxInlineEquation math && ReferenceEquals(math.Equation, paragraph.Equations[0]),
                "The first formula keeps the bookmark's ordered boundary across an ordinary nonpainting run.");
            if (variation.Gap != 0) TestAssert.True(paragraph.Equations[0].SourceRunIndex > anchor.SourceRunIndex,
                "The reader retains the actual raw source gap rather than rewriting source identities.");
        }
    }

    public static void DocxFirstMathBookmarkDoesNotCrossPrefixPicturesOrBoxes()
    {
        foreach (string prefix in new[] { "picture", "box" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var variation in new[] { (Gap: 0, Hidden: false), (Gap: 2, Hidden: false), (Gap: 2, Hidden: true) })
        {
            var paragraph = NonpaintingBookmarkParagraph(ReadNonpaintingBookmark("body", variation.Gap, mode, prefix, variation.Hidden));
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.True(prefix == "picture" ? items.OfType<DocxInlinePicture>().Any() : items.OfType<DocxInlineBox>().Any(),
                "The intervening object is a resolved part of the ordered source.");
            var anchor = paragraph.BookmarkAnchors.Single(a => a.Name == "PrefixMathTarget");
            if (anchor.InlineItemIndex is int index)
                TestAssert.True(index >= 0 && index < items.Count &&
                    (prefix == "picture" ? items[index] is DocxInlinePicture : items[index] is DocxInlineBox),
                    "A bookmark before an intervening object must not acquire the later equation's placement.");
            var afterObject = paragraph.BookmarkAnchors.Single(a => a.Name == "MathAfterObject");
            TestAssert.True(afterObject.InlineItemIndex is int next && next >= 0 && next < items.Count &&
                items[next] is DocxInlineEquation math && ReferenceEquals(math.Equation, paragraph.Equations[0]),
                "The following source boundary still belongs to its own equation.");
        }
    }

    public static void DocxFirstMathBookmarkKeepsInterveningVisibleTextBoundary()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool hidden in new[] { false, true })
        {
            var paragraph = NonpaintingBookmarkParagraph(ReadNonpaintingBookmark("body", 1, mode, "text", hidden));
            var anchor = paragraph.BookmarkAnchors.Single(a => a.Name == "PrefixMathTarget");
            TestAssert.True(anchor.InlineItemIndex is null ||
                paragraph.InlineContent![anchor.InlineItemIndex.Value] is DocxInlineText,
                "A bookmark before visible text must not jump to a later formula.");
        }
    }

    public static void DocxNonpaintingRunsKeepPaintAndMathematicalLinkDestinations()
    {
        foreach (string context in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string prefix in new[] { "none", "text" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = renderer.RenderPagesAsync(ReadNonpaintingBookmark(context, 0, mode, prefix), null, default)
                .ToBlockingEnumerable().ToArray();
            var expected = control.SelectMany(p => p.Annotations).Single(a => a.Destination is not null).Destination!.Value;
            string expectedText = string.Concat(control.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (var variation in new[] { (Gap: 1, Hidden: false), (Gap: 3, Hidden: false),
                (Gap: 1, Hidden: true), (Gap: 3, Hidden: true) })
            {
                var pages = renderer.RenderPagesAsync(ReadNonpaintingBookmark(context, variation.Gap, mode, prefix, variation.Hidden), null, default)
                    .ToBlockingEnumerable().ToArray();
                var actual = pages.SelectMany(p => p.Annotations).Single(a => a.Destination is not null).Destination!.Value;
                TestAssert.Equal(control.Length, pages.Length);
                TestAssert.Equal(expectedText, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(1, pages.Sum(p => p.Images.Count));
                TestAssert.Equal(expected.PageIndex, actual.PageIndex);
                TestAssert.True(Math.Abs(expected.Left!.Value - actual.Left!.Value) < .000001d &&
                    Math.Abs(expected.Top!.Value - actual.Top!.Value) < .000001d,
                    "A nonpainting run changes neither the painted paragraph nor the formula's published destination.");
            }
        }
    }

    private static DocxParagraph NonpaintingBookmarkParagraph(DocxDocument document) =>
        DocxBlockTraversal.EnumerateBodyParagraphs(document)
            .Single(p => p.BookmarkAnchors.Any(a => a.Name == "PrefixMathTarget"));

    private static DocxDocument ReadNonpaintingBookmark(string context, int gap, OoxPdfDocxMarkupMode mode,
        string prefixObject = "none", bool hidden = false) => ReadMutation(prefixObject == "box"
            ? "InlineBoxMath/inline-box-body-equation-inline-box-valid" : "CompositionProbes/body", mode, false, false,
        (xml, w, m) =>
        {
            var paragraph = xml.Root!.Element(w + "body")!.Elements(w + "p")
                .First(p => p.Elements(m + "oMath").Any() || p.Elements(m + "oMathPara").Any());
            var first = paragraph.Elements().First(e => e.Name == m + "oMath" || e.Name == m + "oMathPara");
            XElement? drawing = null;
            if (prefixObject is "picture" or "box")
            {
                drawing = paragraph.Elements(w + "r").First(run => prefixObject == "box"
                    ? run.Descendants(w + "txbxContent").Any()
                    : run.Descendants(w + "drawing").Any() && !run.Descendants(w + "txbxContent").Any());
                drawing.Remove();
            }
            first.AddBeforeSelf(new XElement(w + "r", new XElement(w + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"), "PublicPrefix ")),
                new XElement(w + "bookmarkStart", new XAttribute(w + "id", "907"), new XAttribute(w + "name", "PrefixMathTarget")));
            for (int index = 0; index < gap; index++) first.AddBeforeSelf(hidden
                ? new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "t", "HiddenProbe"))
                : new XElement(w + "r"));
            if (prefixObject == "text") first.AddBeforeSelf(new XElement(w + "r",
                new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Visible prefix ")));
            if (drawing is not null) first.AddBeforeSelf(drawing);
            first.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "907")));
            if (drawing is not null)
            {
                first.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "908"),
                    new XAttribute(w + "name", "MathAfterObject")));
                first.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "908")));
            }
            if (prefixObject != "box")
            {
                var last = paragraph.Elements(m + "oMath").Last();
                last.Descendants(m + "t").Single().Value = "u+v=w";
                last.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PrefixMathTarget"), new XElement(last)));
            }
            paragraph.Add(new XElement(w + "r", new XElement(w + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"), " PublicSuffix")));
            if (context == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
        });
}
