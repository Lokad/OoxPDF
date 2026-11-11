using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrdinaryPictureBookmarksFollowTheirPlacedSourceLine()
    {
        foreach (bool cell in new[] { false, true })
        foreach (bool wrapped in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Read(bool bookmark) => ReadMutation("CompositionProbes/inline-math-navigation", mode,
                false, false, (xml, w, _) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var picture = new XElement(xml.Descendants(w + "drawing").First().Parent!);
                    body.Elements().Where(element => element.Name != w + "sectPr").Remove();
                    xml.Descendants(w + "headerReference").Remove();
                    xml.Descendants(w + "footerReference").Remove();
                    var paragraph = new XElement(w + "p",
                        new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                            string.Concat(Enumerable.Repeat("PublicPrefix ", wrapped ? 12 : 1)))));
                    if (bookmark) paragraph.Add(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "913"),
                        new XAttribute(w + "name", "OrdinaryPictureTarget")));
                    paragraph.Add(picture);
                    if (bookmark) paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "913")));
                    paragraph.Add(new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " PublicSuffix")));
                    var navigation = new XElement(w + "p", new XElement(w + "hyperlink",
                        new XAttribute(w + "anchor", "OrdinaryPictureTarget"), new XElement(w + "r", new XElement(w + "t", "Jump to picture"))));
                    XElement content = cell ? new XElement(w + "tbl",
                        new XElement(w + "tblPr", new XElement(w + "tblW", new XAttribute(w + "w", "8640"), new XAttribute(w + "type", "dxa"))),
                        new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                        new XElement(w + "tr", new XElement(w + "tc", new XElement(w + "tcPr",
                            new XElement(w + "tcW", new XAttribute(w + "w", "8640"), new XAttribute(w + "type", "dxa"))), paragraph))) : paragraph;
                    body.Element(w + "sectPr")!.AddBeforeSelf(content, navigation);
                });
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var source = Read(true);
            var page = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            var control = renderer.RenderPagesAsync(Read(false), null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(control.Content, page.Content);
            var image = page.Images.Single();
            var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" + Regex.Escape(image.ResourceName) + @" Do");
            TestAssert.True(draw.Success, "The bookmark must address an actual placed picture.");
            double Number(int group) => double.Parse(draw.Groups[group].Value, CultureInfo.InvariantCulture);
            var target = page.Annotations.Single(annotation => annotation.IsDestination).Destination!.Value;
            TestAssert.Equal(0, target.PageIndex);
            TestAssert.True(Math.Abs(target.Left - (Number(3) - 3d)) <= .001d,
                "Picture bookmarks retain the existing viewport context at the placed picture, including wrapped body/cell lines.");
            TestAssert.True(Math.Abs(target.Top - (Number(4) + Number(2))) <= .001d,
                "The picture owns the top of its measured source line.");
            string painted = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Select(show => show.Text));
            TestAssert.Equal(wrapped ? 12 : 1, Regex.Matches(painted, "PublicPrefix").Count);
            TestAssert.Contains("PublicSuffix", painted);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }
}
