using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterOuterHiddenCaptionsUseNativeCenteredBounds()
    {
        XDocument source = OuterChart();
        foreach (var label in source.Descendants(CurveC + "tickLblPos")) label.SetAttributeValue("val", "none");
        PdfPage page = RenderEdge(source);
        var axes = EdgeAxes(page);
        // Independent Office16/API/PDF control: source500x432pt, factor(.04,.03),
        // size(.78,.72), no captions. Coordinates below are PDF bottom-origin.
        double[] expectedX = [147d, 119.52d, 537d, 119.52d], expectedY = [147d, 119.52d, 147d, 430.56d];
        TestAssert.True(axes.X.Zip(expectedX).Concat(axes.Y.Zip(expectedY)).All(p => Math.Abs(p.First - p.Second) < .04d),
            "A hidden-caption outer target is the centered source rectangle, without default plot-ratio offsets.");
        TestAssert.Equal("79", EdgeValueLabel(page).Text);
    }

    public static void PptxScatterOuterCaptionsConsumeTheAuthoredArea()
    {
        XDocument source = OuterChart();
        PdfPage page = RenderEdge(source);
        var visible = EdgeAxes(page);
        foreach (var label in source.Descendants(CurveC + "tickLblPos")) label.SetAttributeValue("val", "none");
        var hidden = EdgeAxes(RenderEdge(source));
        TestAssert.True(visible.X[0] > hidden.X[0] && visible.X[2] < hidden.X[2] &&
            visible.Y[1] > hidden.Y[1] && visible.Y[3] < hidden.Y[3], "All four caption bands fit inside the authored outer rectangle.");
        TestAssert.Equal("0,1,2,3,4,5,6,7,8,7,32,57,82", string.Join(",", AutoLabels(page)));
        TestAssert.Equal("79", EdgeValueLabel(page).Text);
    }

    public static void PptxScatterOuterHighCaptionsMirrorPlacementAndPointLabels()
    {
        XDocument source = OuterChart();
        PdfPage lowPage = RenderEdge(source);
        var low = EdgeAxes(lowPage);
        foreach (var label in source.Descendants(CurveC + "tickLblPos")) label.SetAttributeValue("val", "high");
        PdfPage highPage = RenderEdge(source);
        var high = EdgeAxes(highPage);
        Near(684d, low.X[0] + high.X[2]);
        Near(550.08d, low.Y[1] + high.Y[3]);
        var before = EdgeValueLabel(lowPage); var after = EdgeValueLabel(highPage);
        Near(high.X[0] - low.X[0], after.X - before.X);
        Near(high.Y[1] - low.Y[1], after.Y - before.Y);
        TestAssert.Equal(string.Join(",", AutoLabels(lowPage)), string.Join(",", AutoLabels(highPage)));
    }

    public static void PptxScatterOuterEquivalentPositionsPreserveEveryOperation()
    {
        XDocument source = OuterChart();
        string factor = RenderEdge(source).Content;
        source.Descendants(CurveC + "xMode").Single().SetAttributeValue("val", "edge");
        source.Descendants(CurveC + "x").Single().SetAttributeValue("val", ".15");
        source.Descendants(CurveC + "yMode").Single().SetAttributeValue("val", "edge");
        source.Descendants(CurveC + "y").Single().SetAttributeValue("val", ".17");
        TestAssert.Equal(factor, RenderEdge(source).Content);
    }

    public static void PptxScatterOuterUnmeasuredSizeModesKeepTheirFallback()
    {
        foreach (string axis in new[] { "w", "h" })
        {
            XDocument source = OuterChart();
            source.Descendants(CurveC + (axis + "Mode")).Single().SetAttributeValue("val", "edge");
            string actual = RenderEdge(source).Content;
            source.Descendants(CurveC + "layoutTarget").Single().SetAttributeValue("val", "unresolved");
            TestAssert.Equal(RenderEdge(source).Content, actual);
        }
    }

    private static XDocument OuterChart()
    {
        XDocument source = EdgeChart(.04d, .78d, "factor");
        source.Descendants(CurveC + "layoutTarget").Single().SetAttributeValue("val", "outer");
        return source;
    }
}
