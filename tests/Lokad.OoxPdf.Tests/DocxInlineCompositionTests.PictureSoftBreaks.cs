using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrdinaryPictureSoftBreaksKeepFollowingWordsClearOfPaint()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (bool mathFont in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int breaks in new[] { 0, 1, 2, 3 })
        foreach (bool merge in new[] { false, true })
        {
            var source = ReadMutation("docx-picture-rotated-raster", mode, false, false,
                (xml, w, _) =>
                {
                    var paragraph = xml.Descendants(w + "p").Single();
                    var drawing = paragraph.Descendants(w + "drawing").Single();
                    var transform = drawing.Descendants(a + "xfrm").Single();
                    transform.Attribute("rot")!.Remove();
                    transform.Element(a + "ext")!.SetAttributeValue("cx", "914400");
                    transform.Element(a + "ext")!.SetAttributeValue("cy", "1828800");
                    var first = paragraph.Elements(w + "r").First();
                    for (int index = 0; index < breaks; index++) first.Add(new XElement(w + "br"));
                    if (merge) MergePictureRuns(paragraph, w);
                });
            var page = PictureRunBoundaryRenderer(mode, mathFont).RenderPagesAsync(source, null, default)
                .ToBlockingEnumerable().Single();
            var draw = Regex.Match(page.Content,
                @"(?<Width>[-\d.]+) 0 0 (?<Height>[-\d.]+) (?<X>[-\d.]+) (?<Y>[-\d.]+) cm\s+/Im\d+ Do");
            TestAssert.True(draw.Success, "The source picture must actually paint.");
            double right = double.Parse(draw.Groups["X"].Value, CultureInfo.InvariantCulture) +
                double.Parse(draw.Groups["Width"].Value, CultureInfo.InvariantCulture);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            TestAssert.Equal(1, shows.Count(show => show.Text.Trim() == "Before"));
            var after = shows.Single(show => show.Text.Trim() == "After");
            TestAssert.True(after.X >= right - .001d,
                "The following word must remain visible beside the opaque picture after supported soft breaks.");
        }
    }
}
