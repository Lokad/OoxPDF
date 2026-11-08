using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxLabelSetDeletionKeepsRawStateAndDefinedChildDefaults()
    {
        foreach (string? raw in new string?[] { null, "", "1", "true", "0", "false" })
        {
            string deletion = raw is null ? "" : raw.Length == 0 ? "<c:delete/>" : $"<c:delete val=\"{raw}\"/>";
            var (chart, plot) = Source("<c:dLbls>" + deletion + "</c:dLbls>");
            bool? expected = raw is null ? null : raw is not ("0" or "false");
            TestAssert.Equal(expected, chart.Plots[0].DataLabels.IsDeleted);
            TestAssert.Equal(raw ?? "", chart.Plots[0].DataLabels.IsDeletedValue);
            TestAssert.Equal(null, chart.Plots[0].DataLabels.ShowValue);
            TestAssert.Equal(expected != true, Flag(Invoke("ResolveChartDataLabelOptions", Read(chart), 1), "HasVisibleContent"));
        }

        const string pointRequest = "<c:dLbl><c:idx val=\"0\"/><c:showVal val=\"1\"/></c:dLbl>";
        var (parent, parentPlot) = Source("<c:dLbls>" + pointRequest + "<c:delete val=\"1\"/></c:dLbls>");
        TestAssert.Equal(null, parent.Plots[0].DataLabels.Overrides[0].ShowSeriesName);
        object parentOptions = Read(parent);
        object point = Invoke("ResolveChartDataLabelOptions", parentOptions, 0);
        foreach (var component in Components)
            TestAssert.True(Flag(point, component.Property), "A defined point owns its defaults inside a deleted label set.");
        TestAssert.True(!Flag(Invoke("ResolveChartDataLabelOptions", parentOptions, 1), "HasVisibleContent"),
            "The point request does not reactivate neighboring default labels.");
        var (series, seriesPlot) = Source("<c:dLbls><c:delete val=\"1\"/></c:dLbls>", "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>");
        object options = Read(series);
        object children = Invoke("ReadSceneSeriesDataLabelOptions", series,
            series.Plots[0]);
        object merged = Invoke("ResolveChartDataLabelOptionsForSeries", options, children, 0);
        TestAssert.True(Flag(Invoke("ResolveChartDataLabelOptions", merged, 1), "HasVisibleContent"),
            "A defined series owns its missing deletion default, independently of a deleted plot.");
    }
}
