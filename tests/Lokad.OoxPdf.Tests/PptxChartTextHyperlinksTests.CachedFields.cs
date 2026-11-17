using System.Xml.Linq;
using Lokad.OoxPdf.Pptx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartTextHyperlinksTests
{
    public static void PptxChartCachedFieldsKeepOrderedTextStyleAndLinkOwners()
    {
        foreach (bool inheritParagraphFormatting in new[] { false, true })
        foreach (bool field in new[] { false, true })
        foreach (bool mixed in new[] { false, true })
        {
            string run = CachedFieldInline(field, mixed);
            var text = XElement.Parse($$"""
                <c:tx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="2400" b="1"/></a:pPr>{{run}}</a:p></c:rich>
                </c:tx>
                """);
            string expected = mixed ? "STARTCACHEDEND" : "CACHED";
            var runs = PptxSceneBuilder.ReadChartTextRuns(text, PptxTheme.Empty, PptxColorMap.Default, inheritParagraphFormatting);
            TestAssert.Equal(expected, PptxSceneBuilder.ReadChartText(text, false)!);
            TestAssert.Equal(expected, string.Concat(runs.Select(item => item.Text)));
            var cached = runs.Single(item => item.Text == "CACHED");
            TestAssert.Equal(18d, cached.TextStyle.FontSize!.Value);
            TestAssert.Equal("Arial", cached.TextStyle.FontFamily!);
            TestAssert.Equal(new RgbColor(0, 0, 255), cached.TextStyle.Color!.Value);
            TestAssert.Equal(false, cached.TextStyle.Bold!.Value);
            TestAssert.Equal("rIdField", cached.HyperlinkClickId!);
        }
    }

    internal static string CachedFieldInline(bool field, bool mixed)
    {
        string tag = field ? "fld" : "r";
        string attributes = field ? " id=\"{F3AC8070-99BD-4B86-B07F-E4F81F829748}\" type=\"datetime1\"" : "";
        string cached = $"<a:{tag}{attributes}><a:rPr sz=\"1800\" b=\"0\"><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill><a:latin typeface=\"Arial\"/><a:hlinkClick r:id=\"rIdField\"/></a:rPr><a:t>CACHED</a:t></a:{tag}>";
        return mixed ? "<a:r><a:rPr b=\"0\"/><a:t>START</a:t></a:r>" + cached + "<a:r><a:rPr b=\"0\"/><a:t>END</a:t></a:r>" : cached;
    }
}
