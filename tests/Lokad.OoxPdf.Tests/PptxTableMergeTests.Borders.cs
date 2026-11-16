using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxTableMergeTests
{
    public static void PptxTableBordersKeepDisconnectedSourceIntervals()
    {
        foreach (bool vertical in new[] { false, true })
        {
            PdfPage page = RenderBorderIntervals(vertical, null);
            string[] expected = vertical
                ? ["180 252 m 180 324.375 l S", "180 395.625 m 180 468 l S"]
                : ["72 468 m 180.375 468 l S", "287.625 468 m 396 468 l S"];
            TestAssert.True(BorderStrokes(page).SequenceEqual(expected),
                "An explicitly borderless middle cell separates the two source edges at the default 0.75pt width: " + string.Join("; ", BorderStrokes(page)));
            CheckBorderContent(page);
        }
    }

    public static void PptxTableBordersMergeConnectedEdgesAndKeepStyleOwners()
    {
        foreach (bool vertical in new[] { false, true })
        {
            PdfPage connected = RenderBorderIntervals(vertical, "FF0000");
            TestAssert.Equal(vertical ? "180 252 m 180 468 l S" : "72 468 m 396 468 l S", BorderStrokes(connected).Single());
            CheckBorderContent(connected);
            PdfPage distinct = RenderBorderIntervals(vertical, "0000FF");
            TestAssert.Equal(3, BorderStrokes(distinct).Length);
            TestAssert.Contains("0 0 1 RG", distinct.Content);
            TestAssert.True(!BorderStrokes(distinct).Contains(vertical ? "180 252 m 180 468 l S" : "72 468 m 396 468 l S"),
                "A differently styled middle edge does not reconnect the separated red intervals.");
            CheckBorderContent(distinct);
        }
    }

    private static string[] BorderStrokes(PdfPage page) => Regex.Matches(page.Content, @"(?m)^[^\r\n]+ m [^\r\n]+ l S\r?$")
        .Select(match => match.Value.TrimEnd('\r')).ToArray();

    private static void CheckBorderContent(PdfPage page)
    {
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
        foreach (string label in new[] { "ONE", "GAP", "THREE" }) TestAssert.Contains(label, text);
        TestAssert.True(page.ExtGStates.Any(state => state.StrokeAlpha == .5d), "Union keeps the authored border opacity.");
        PdfContentValidator.ValidatePage(page, 0, default);
    }

    private static PdfPage RenderBorderIntervals(bool vertical, string? middleColor)
    {
        string Cell(int index)
        {
            string? color = index == 1 ? middleColor : "FF0000";
            string border = string.Concat(new[] { "L", "R", "T", "B" }.Select(side =>
                "<a:ln" + side + ">" + (color is not null && side == (vertical ? "R" : "T")
                    ? "<a:solidFill><a:srgbClr val=\"" + color + "\"><a:alpha val=\"50000\"/></a:srgbClr></a:solidFill>"
                    : "<a:noFill/>") + "</a:ln" + side + ">"));
            string text = new[] { "ONE", "GAP", "THREE" }[index];
            return "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"1800\"><a:latin typeface=\"DejaVu Sans\"/></a:rPr><a:t>" +
                text + "</a:t></a:r></a:p></a:txBody><a:tcPr>" + border + "</a:tcPr></a:tc>";
        }
        string grid = string.Concat(Enumerable.Repeat("<a:gridCol w=\"1371600\"/>", vertical ? 1 : 3));
        string rows = vertical ? string.Concat(Enumerable.Range(0, 3).Select(index => "<a:tr h=\"914400\">" + Cell(index) + "</a:tr>"))
            : "<a:tr h=\"914400\">" + string.Concat(Enumerable.Range(0, 3).Select(Cell)) + "</a:tr>";
        string frame = $"<p:graphicFrame><p:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"{(vertical ? 1371600 : 4114800)}\" cy=\"{(vertical ? 2743200 : 914400)}\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblGrid>{grid}</a:tblGrid>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(), ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(), ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/slides/slide1.xml"] = $"<p:sld xmlns:p=\"{P}\" xmlns:a=\"{A}\"><p:cSld><p:spTree>{frame}</p:spTree></p:cSld></p:sld>"
        });
        using var package = OoxPackage.Open(input, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
