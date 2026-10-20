using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlinePictureLinksFollowPaintedFramesAndPreserveParentContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent picture origins use installed Windows fonts.");
        CheckInlinePictureLinks(null, compareWord: true);
    }

    internal static void CheckInlinePictureLinks(IFontResolver? resolver, bool compareWord)
    {
        using var gold = ReadInlineBoxMathGold("inline-picture-links-office.json");
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var source = ReadMutation("InlineBoxMath/" + row.GetProperty("Id").GetString(), mode, false, false);
            var parent = ReadMutation("InlineBoxMath/" + Path.GetFileNameWithoutExtension(row.GetProperty("SourceInput").GetString()), mode, false, false);
            var page = renderer.RenderBlankPages(source, null, default).Single();
            var before = renderer.RenderBlankPages(parent, null, default).Single();
            TestAssert.Equal(before.Content, page.Content);
            TestAssert.Equal(before.Images.Count, page.Images.Count);
            var image = page.Images.Single();
            var links = page.Annotations.Where(link => link.Uri == row.GetProperty("TargetUri").GetString()).ToArray();
            TestAssert.Equal(row.GetProperty("ExpectedUriCount").GetInt32(), links.Length);
            if (links.Length == 0) continue;
            var link = links.Single();
            var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" + Regex.Escape(image.ResourceName) + @" Do");
            TestAssert.True(draw.Success, "The clickable region belongs to an actual axis-aligned raster draw.");
            double Number(int index) => double.Parse(draw.Groups[index].Value, CultureInfo.InvariantCulture);
            TestAssert.True(Math.Abs(link.X - Number(3)) <= .001d, "Link X follows the painted frame.");
            TestAssert.True(Math.Abs(link.Y - Number(4)) <= .001d, "Link Y follows the painted frame.");
            TestAssert.True(Math.Abs(link.Width - Number(1)) <= .001d, "Link width follows the painted frame.");
            TestAssert.True(Math.Abs(link.Height - Number(2)) <= .001d, "Link height follows the painted frame.");
            if (compareWord && row.GetProperty("WordLinkGeometryAdmitted").GetBoolean())
            {
                var rect = row.GetProperty("WordLinks")[0].GetProperty("Rect");
                double[] actual = [link.X, page.Height - link.Y - link.Height, link.X + link.Width, page.Height - link.Y];
                for (int index = 0; index < actual.Length; index++)
                    TestAssert.True(Math.Abs(actual[index] - rect[index].GetDouble()) <= 1d, "Final URI bounds retain the original Word geometry gate.");
            }
        }
    }

    public static void DocxInlinePictureLinksKeepUnsupportedActionsAndTransformsUnlinked()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        foreach (string variant in new[] { "missing", "action", "hover", "rotation" })
        {
            var source = ReadMutation("InlineBoxMath/inline-picture-link-properties-body", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, _, _) =>
                {
                    var click = xml.Descendants(a + "hlinkClick").Single();
                    if (variant == "missing") click.SetAttributeValue(r + "id", "missing");
                    if (variant == "action") click.SetAttributeValue("action", "ppaction://hlinkshowjump?jump=nextslide");
                    if (variant == "hover") click.Name = a + "hlinkHover";
                    if (variant == "rotation") xml.Descendants(a + "xfrm").First().SetAttributeValue("rot", "5400000");
                });
            var renderer = new DocxRenderer(new MapFontResolver([], "Courier"), OoxPdfDocxMarkupMode.Final,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var page = renderer.RenderBlankPages(source, null, default).Single();
            TestAssert.Equal(1, page.Images.Count);
            TestAssert.Equal(0, page.Annotations.Count(link => link.IsUri));
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(source, null, cancellation.Token).ToArray());
        }
    }

    public static void DocxInlinePictureLinksRequireDecodedAndUnclippedPaint()
    {
        var source = ReadMutation("InlineBoxMath/inline-picture-link-properties-body", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = source.Paragraphs.Single(paragraph => paragraph.Images.Count != 0);
        var broken = paragraph.Images.Single() with { Bytes = [0, 1, 2, 3] };
        var replacement = paragraph with { Images = [broken], InlineContent = TestAssert.NotNull(paragraph.InlineContent)
            .Select(item => item is DocxInlinePicture ? new DocxInlinePicture(broken) : item).ToArray() };
        var document = source with { BodyElements = source.BodyElements.Select(item =>
            item is DocxParagraphElement p && ReferenceEquals(p.Paragraph, paragraph) ? new DocxParagraphElement(replacement) : item).ToArray(),
            FallbackParagraphs = source.FallbackParagraphs.Select(p => ReferenceEquals(p, paragraph) ? replacement : p).ToArray() };
        var diagnostics = new List<OoxPdfDiagnostic>();
        var renderer = new DocxRenderer(new MapFontResolver([], "Courier"), OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var page = renderer.RenderBlankPages(document, diagnostics.Add, default).Single();
        TestAssert.Equal(0, page.Images.Count);
        TestAssert.Equal(0, page.Annotations.Count(link => link.IsUri));
        TestAssert.Equal(1, diagnostics.Count(diagnostic => diagnostic.Id == "IMAGE_UNSUPPORTED_FORMAT"));

        var clipped = ReadMutation("InlineBoxMath/inline-picture-link-properties-cell", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, _) =>
            {
                var row = xml.Descendants(w + "tr").Single();
                var properties = row.Element(w + "trPr");
                if (properties is null) { properties = new XElement(w + "trPr"); row.AddFirst(properties); }
                properties.Elements(w + "trHeight").Remove();
                properties.Add(new XElement(w + "trHeight", new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "exact")));
            });
        var clippedPage = renderer.RenderBlankPages(clipped, null, default).Single();
        TestAssert.Equal(1, clippedPage.Images.Count);
        TestAssert.Equal(0, clippedPage.Annotations.Count(link => link.IsUri));
    }
}
