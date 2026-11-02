using System.Xml.Linq;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxPointTextStylesOwnChartDefaultsAcrossKinds()
    {
        foreach (string kind in new[] { "barChart", "lineChart", "areaChart", "scatterChart", "bubbleChart", "pieChart", "doughnutChart" })
        foreach (bool sparseRoot in new[] { false, true })
        foreach (bool scene in new[] { false, true })
        {
            var (first, firstPlot) = TextScopeSource(kind, TextScopeParent, "", sparseRoot: sparseRoot);
            var (other, otherPlot) = TextScopeSource(kind, TextScopeOtherParent, "", sparseRoot: sparseRoot);
            var (plain, plainPlot) = TextScopeSource(kind, "", "", sparseRoot: sparseRoot);
            object Point(PptxSceneChart chart, XElement plot) => Invoke("ResolveChartDataLabelOptions", Read(chart, plot, scene), 2);
            object firstStyle = TextScopeStyle(Point(first, firstPlot));
            TestAssert.Equal(TextScopeStyle(Point(plain, plainPlot)), firstStyle);
            TestAssert.Equal(firstStyle, TextScopeStyle(Point(other, otherPlot)));
            TestAssert.True(!TextScopeStyle(Invoke("ResolveChartDataLabelOptions", Read(first, firstPlot, scene), 0))
                .Equals(TextScopeStyle(Invoke("ResolveChartDataLabelOptions", Read(other, otherPlot, scene), 0))),
                "Changing container style still changes an unoverridden label.");
            TestAssert.Equal(null, first.Plots[0].DataLabels.Overrides.Single().TextStyle.FontFamily);
            TestAssert.Equal(null, first.Plots[0].DataLabels.Overrides.Single().TextStyle.FontSize);
        }
    }

    public static void PptxDefinedSeriesTextStylesOwnDefaults()
    {
        foreach (bool scene in new[] { false, true })
        foreach (bool defined in new[] { false, true })
        {
            var (chart, plot) = TextScopeSource("lineChart", TextScopeParent, "", series: defined ? "<c:dLbls/>" : "", sparseRoot: true);
            object original = Read(chart, plot, scene);
            object perSeries = Invoke("ReadSceneOrXmlSeriesDataLabelOptions", scene ? chart : null, scene ? chart.Plots[0] : null,
                plot, PptxTheme.Empty, PptxColorMap.Default);
            object resolved = Invoke("ResolveChartDataLabelOptionsForSeries", original, perSeries, 0);
            var (baseline, baselinePlot) = TextScopeSource("lineChart", "", "", sparseRoot: true);
            TestAssert.Equal(defined ? TextScopeStyle(Read(baseline, baselinePlot, scene)) : TextScopeStyle(original), TextScopeStyle(resolved));
        }
    }

    public static void PptxPointTextStylesPreserveOverridesAndProvenance()
    {
        const string own = """
            <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200" b="0" i="0" u="none" strike="noStrike"><a:latin typeface="Point Face"/></a:defRPr></a:pPr></a:p></c:txPr>
            """;
        var (chart, plot) = TextScopeSource("lineChart", TextScopeParent, own);
        var raw = chart.Plots[0].DataLabels.Overrides.Single().TextStyle;
        foreach (object options in ReadBoth(chart, plot))
        {
            object style = TextScopeStyle(Invoke("ResolveChartDataLabelOptions", options, 2));
            TestAssert.Equal("Point Face", TextScopeProperty(style, "FontFamily"));
            TestAssert.Equal(12d, TextScopeProperty(style, "FontSize"));
            TestAssert.Equal(.5d, TextScopeProperty(style, "CharacterSpacing"));
            TestAssert.Equal(new RgbColor(1, 2, 3), TextScopeProperty(style, "Color"));
            TestAssert.Equal(.8d, TextScopeProperty(style, "Alpha"));
            foreach (string flag in new[] { "Bold", "Italic", "Underline", "Strike" }) TestAssert.Equal(false, TextScopeProperty(style, flag));
            TestAssert.Equal(raw.RequestedTypeface, TextScopeProperty(style, "RequestedTypeface"));
            TestAssert.Equal(raw.TypefaceSource, TextScopeProperty(style, "TypefaceSource"));
        }
        TestAssert.Equal(raw, chart.Plots[0].DataLabels.Overrides.Single().TextStyle);
        TestAssert.Equal(null, raw.Color); TestAssert.Equal(null, raw.Alpha); TestAssert.Equal(null, raw.CharacterSpacing);
    }

    public static void PptxPointTextRenderingIgnoresUnusedPlotStyle()
    {
        string Render(string style)
        {
            var xml = ManualFrameSource(SizedFrameLayout(), false);
            xml.Descendants(C + "dLbls").Single().Add(XElement.Parse("<c:holder xmlns:c=\"" + C + "\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" + style + "</c:holder>").Elements());
            return RenderChartPage(xml).Content;
        }
        TestAssert.Equal(Render(""), Render(TextScopeParent));
        TestAssert.Equal(Render(TextScopeParent), Render(TextScopeOtherParent));
    }

    private const string TextScopeParent = """
        <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="2000" b="1" i="1" u="sng" strike="sngStrike" spc="150"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="30000"/></a:srgbClr></a:solidFill><a:latin typeface="Plot Face"/></a:defRPr></a:pPr></a:p></c:txPr>
        """;
    private const string TextScopeOtherParent = """
        <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="3000" b="0"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:latin typeface="Other Plot Face"/></a:defRPr></a:pPr></a:p></c:txPr>
        """;

    private static (PptxSceneChart Chart, XElement Plot) TextScopeSource(string kind, string parent, string point, string series = "", bool sparseRoot = false)
    {
        var (_, plot) = Source("<c:dLbls><c:dLbl><c:idx val=\"2\"/>" + point + "</c:dLbl><c:showVal val=\"1\"/>" + parent + "</c:dLbls>", series);
        plot.Name = C + kind;
        string rootStyle = sparseRoot ? "<c:txPr><a:p><a:pPr><a:defRPr sz=\"1000\"/></a:pPr></a:p></c:txPr>" : """
            <c:txPr><a:p><a:pPr><a:defRPr sz="1000" b="1" i="1" u="sng" strike="sngStrike" spc="50"><a:solidFill><a:srgbClr val="010203"><a:alpha val="80000"/></a:srgbClr></a:solidFill><a:latin typeface="Root Face"/></a:defRPr></a:pPr></a:p></c:txPr>
            """;
        plot.Document!.Root!.Element(C + "txPr")!.ReplaceWith(XElement.Parse("<c:holder xmlns:c=\"" + C + "\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" + rootStyle + "</c:holder>").Elements());
        return (PptxTests.BuildSingleChartScene(plot.Document.ToString())!, plot);
    }

    private static object TextScopeStyle(object options) => TextScopeProperty(options, "TextStyle")!;
    private static object? TextScopeProperty(object value, string property) => value.GetType().GetProperty(property)!.GetValue(value);
}
