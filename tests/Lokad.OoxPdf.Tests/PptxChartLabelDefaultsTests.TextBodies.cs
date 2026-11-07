using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxPointTextBodyDefaultsIgnoreUnownedFormatting()
    {
        foreach (string kind in new[] { "barChart", "lineChart", "areaChart", "scatterChart", "bubbleChart", "pieChart", "doughnutChart" })
        foreach (string paragraph in new[] { "", "<a:pPr/>", "<a:endParaRPr sz=\"1400\"/>", "<a:pPr/><a:endParaRPr sz=\"1400\"/>" })
        foreach (bool list in new[] { false, true })
        foreach (bool sparse in new[] { false, true })
        {
            string body = TextBodySource(paragraph, list);
            var (chart, plot) = TextScopeSource(kind, TextScopeParent, body, sparseRoot: sparse);
            var raw = chart.Plots[0].DataLabels.Overrides.Single().TextStyle;
            TestAssert.Equal(null, raw.FontFamily); TestAssert.Equal(null, raw.FontSize);
            foreach (object options in ReadBoth(chart, plot))
            {
                object style = TextScopeStyle(Invoke("ResolveChartDataLabelOptions", options, 2));
                TestAssert.Equal("Arial", TextScopeProperty(style, "FontFamily"));
                TestAssert.Equal(18d, TextScopeProperty(style, "FontSize"));
                TestAssert.Equal(0d, TextScopeProperty(style, "CharacterSpacing"));
                TestAssert.Equal(new RgbColor(0, 0, 0), TextScopeProperty(style, "Color"));
                TestAssert.Equal(1d, TextScopeProperty(style, "Alpha"));
                foreach (string flag in new[] { "Bold", "Italic", "Underline", "Strike" })
                    TestAssert.Equal(false, TextScopeProperty(style, flag));
                TestAssert.True(TextScopeProperty(TextScopeStyle(options), "FontSize") is 20d,
                    "The point's implicit body cannot replace its container or neighboring points.");
            }
        }
    }

    public static void PptxPointTextBodyAuthoredDefaultsPreserveChartFormatting()
    {
        foreach (string kind in new[] { "barChart", "lineChart" })
        foreach (bool list in new[] { false, true })
        foreach (string defaults in new[] { "<a:defRPr/>", "<a:defRPr sz=\"1200\"/>" })
        {
            var (chart, plot) = TextScopeSource(kind, TextScopeParent,
                TextBodySource("<a:pPr>" + defaults + "</a:pPr>", list));
            foreach (object options in ReadBoth(chart, plot))
            {
                object style = TextScopeStyle(Invoke("ResolveChartDataLabelOptions", options, 2));
                TestAssert.Equal("Root Face", TextScopeProperty(style, "FontFamily"));
                TestAssert.Equal(defaults.Contains("1200", StringComparison.Ordinal) ? 12d : 10d,
                    TextScopeProperty(style, "FontSize"));
                TestAssert.Equal(.5d, TextScopeProperty(style, "CharacterSpacing"));
                TestAssert.Equal(new RgbColor(1, 2, 3), TextScopeProperty(style, "Color"));
                TestAssert.Equal(true, TextScopeProperty(style, "Bold"));
            }
        }
    }

    public static void PptxPointTextBodyLineKeyKeepsIndependentParagraphMetrics()
    {
        PdfPage Render(bool authored)
        {
            var xml = KeyPlacementSource("r", 12, true);
            var point = new XElement(C + "dLbl", new XElement(C + "idx", new XAttribute("val", 2)));
            point.Add(XElement.Parse("<c:holder xmlns:c=\"" + C + "\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                TextBodySource(authored ? "<a:pPr><a:defRPr sz=\"1800\"><a:latin typeface=\"Arial\"/></a:defRPr></a:pPr>" : "", false) + "</c:holder>").Elements());
            xml.Descendants(C + "dLbls").Single().AddFirst(point);
            return RenderChartPage(xml);
        }
        var implicitPage = Render(false); var authoredPage = Render(true);
        var implicitShows = LineParagraphShows(implicitPage); var authoredShows = LineParagraphShows(authoredPage);
        TestAssert.Equal("3", string.Concat(implicitShows.Select(s => s.Text)));
        TestAssert.Equal("3", string.Concat(authoredShows.Select(s => s.Text)));
        TestAssert.True(implicitShows.Concat(authoredShows).All(s => Math.Abs(s.Size - 18d) < .002),
            "Implicit and explicit fallback bodies paint the same visible glyph size.");
        var first = LastKeySegment(implicitPage); var second = LastKeySegment(authoredPage);
        TestAssert.True(Math.Abs(first.X2 - first.X1) < .002 && second.X2 - second.X1 > 1d,
            "Independent Office controls distinguish a zero-advance implicit key from an authored paragraph key.");
        TestAssert.True(Math.Abs(implicitShows[0].Y - authoredShows[0].Y) < .002,
            "Key capacity does not move the visible text baseline.");
        PdfContentValidator.ValidatePage(implicitPage, 0, default);
        PdfContentValidator.ValidatePage(authoredPage, 0, default);
    }

    private static string TextBodySource(string paragraph, bool list) => "<c:txPr><a:bodyPr/>" +
        (list ? "<a:lstStyle><a:lvl1pPr><a:defRPr sz=\"1400\"><a:latin typeface=\"List Face\"/></a:defRPr></a:lvl1pPr></a:lstStyle>" : "") +
        "<a:p>" + paragraph + "</a:p></c:txPr>";

    private static (double X1, double X2) LastKeySegment(PdfPage page)
    {
        var matches = Regex.Matches(page.Content, @"(?<x1>[-\d.]+) [-\d.]+ m\s+(?<x2>[-\d.]+) [-\d.]+ l\s+S");
        TestAssert.True(matches.Count > 0, "Expected a painted line-key segment.");
        var last = matches[^1];
        return (double.Parse(last.Groups["x1"].Value, CultureInfo.InvariantCulture),
            double.Parse(last.Groups["x2"].Value, CultureInfo.InvariantCulture));
    }
}
