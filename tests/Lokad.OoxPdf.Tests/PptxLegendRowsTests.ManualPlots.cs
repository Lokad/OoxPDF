using System.Text.RegularExpressions;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxManualPlotWidthDoesNotConstrainTheBottomLegend()
    {
        foreach (string kind in new[] { "column", "bar" })
        {
            var narrow = Render(576, 18, 2, overlay: true, kind: kind, plotLayout: ManualPlot("0.25"));
            var wide = Render(576, 18, 2, overlay: true, kind: kind, plotLayout: ManualPlot("0.65"));
            var narrowNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(narrow.Page)
                .Where(show => show.Text.Contains("North West Region", StringComparison.Ordinal)).ToArray();
            var wideNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(wide.Page)
                .Where(show => show.Text.Contains("North West Region", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(2, narrowNames.Sum(show => Regex.Matches(show.Text, @"North West Region \d+").Count));
            TestAssert.True(narrowNames.Select(show => (show.Text, show.X, show.Y))
                    .SequenceEqual(wideNames.Select(show => (show.Text, show.X, show.Y))),
                "Changing authored data width must preserve the default legend's chart-frame placement.");
        }
    }

    public static void PptxManualPlotLegendsWrapWithoutReservingTheAuthoredDataFrame()
    {
        foreach (string kind in new[] { "column", "bar" })
        foreach (int width in new[] { 288, 576 })
        {
            var overlay = Render(width, 12, 6, overlay: true, kind: kind, plotLayout: ManualPlot("0.25"));
            var reserved = Render(width, 12, 6, overlay: false, kind: kind, plotLayout: ManualPlot("0.25"));
            TestAssert.True(CheckCompleteNamesInFrame(overlay.Page, width, 6) > 1,
                "Six complete names wrap within the chart rather than following the narrower data plot.");
            CheckCompleteNamesInFrame(reserved.Page, width, 6);
            string[] Rectangles(string content) => Regex.Matches(content, @"[-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ re f")
                .Select(match => match.Value).ToArray();
            TestAssert.True(Rectangles(overlay.Page.Content).SequenceEqual(Rectangles(reserved.Page.Content)),
                "Legend overlay changes must not resize or displace authored bars or legend keys.");
            TestAssert.True(!overlay.Diagnostics.Concat(reserved.Diagnostics)
                    .Any(item => item.Message.Contains("overflowing entries", StringComparison.Ordinal)),
                "All six supported names fit the available chart-frame rows.");
        }
    }

    private static string ManualPlot(string width) => $$"""
        <c:layout><c:manualLayout><c:layoutTarget val="inner"/>
          <c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/>
          <c:x val="0.15"/><c:y val="0.2"/><c:w val="{{width}}"/><c:h val="0.55"/>
        </c:manualLayout></c:layout>
        """;
}
