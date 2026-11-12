using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    private const string AdmissionPlot = "<c:layout><c:manualLayout><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/>" +
        "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".1\"/><c:y val=\".1\"/>" +
        "<c:w val=\".6\"/><c:h val=\".7\"/></c:manualLayout></c:layout>";
    private const string AdmissionFlags = "<c:showVal val=\"1\"/><c:showCatName val=\"1\"/>" +
        "<c:showSerName val=\"1\"/><c:showPercent val=\"1\"/><c:showBubbleSize val=\"0\"/>";

    public static void PptxPolarLabelRowsKeepCompleteLabelUnitsApart()
    {
        foreach (string kind in new[] { "pie", "doughnut" })
        foreach (int count in new[] { 3, 6 })
        foreach (bool key in new[] { false, true })
        foreach (bool rich in new[] { false, true })
        foreach (bool math in new[] { false, true })
        {
            string percent = (100d / count).ToString("0.#", CultureInfo.InvariantCulture) + "%";
            string components = "<c:showLegendKey val=\"" + (key ? 1 : 0) + "\"/>" + AdmissionFlags;
            const string background = "<c:spPr><a:solidFill><a:srgbClr val=\"FFF2CC\"/></a:solidFill></c:spPr>";
            string points = rich ? string.Concat(Enumerable.Range(1, count).Select(i =>
                $"<c:dLbl><c:idx val=\"{i - 1}\"/><c:tx><c:rich><a:bodyPr/><a:p>" +
                "<a:r><a:t>Series 1, </a:t></a:r>" +
                $"<a:r><a:rPr><a:hlinkClick r:id=\"rLabel{i}\"/></a:rPr><a:t>Category {i}, 4, {percent}</a:t></a:r>" +
                "</a:p></c:rich></c:tx>" + background + components + "</c:dLbl>")) : "";
            string labels = "<c:dLbls>" + points + background + components + "</c:dLbls>";
            string fills = string.Concat(Enumerable.Range(0, count).Select(i =>
                $"<c:dPt><c:idx val=\"{i}\"/><c:spPr><a:solidFill><a:srgbClr val=\"FF00FE\"/></a:solidFill></c:spPr></c:dPt>"));
            string relations = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                string.Concat(Enumerable.Range(1, count).Select(i => $"<Relationship Id=\"rLabel{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.invalid/label/{i}\" TargetMode=\"External\"/>")) + "</Relationships>";
            var result = Render(576, 12, count, false, kind: kind, categoryLegend: true, plotLayout: AdmissionPlot,
                dataLabels: labels, seriesOptions: "<c:tx><c:v>Series 1</c:v></c:tx><c:explosion val=\"20\"/>" + fills,
                resolver: new PortableRenderingTests.PinnedResolver(math), chartRelationships: relations);
            var ink = new List<(string Text, PdfRectangle Ink)>();
            AssertPolarLabelGlyphInk(result.Page, $"{kind}/{count}/{key}/{rich}/{math}", ink);
            var groups = new List<PdfRectangle>();
            var baselines = new List<double>();
            int cursor = 0;
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(result.Page, false,
                fontResourceFilter: name => name.StartsWith("CP", StringComparison.Ordinal)).ToArray();
            for (int index = 1; index <= count; index++)
            {
                string expected = $"Series 1, Category {index}, 4, {percent}", actual = "";
                PdfRectangle? frame = null;
                int first = cursor;
                while (actual != expected)
                {
                    TestAssert.True(cursor < ink.Count, "Every source label must retain all of its painted components.");
                    var part = ink[cursor++]; actual += part.Text;
                    TestAssert.True(expected.StartsWith(actual, StringComparison.Ordinal), "Label text keeps its source order and punctuation.");
                    frame = frame is { } previous ? UnionLabelInk(previous, part.Ink) : part.Ink;
                }
                groups.Add(frame!.Value); baselines.Add(shows[first].Y);
            }
            TestAssert.Equal(ink.Count, cursor);
            for (int index = 0; index < groups.Count; index++)
                for (int other = 0; other < index; other++)
                    TestAssert.True(!LabelInkOverlaps(groups[index], groups[other]),
                        "Distinct automatic polar labels must not obscure each other's measured glyph ink.");
            var backgrounds = PptxChartLabelDefaultsTests.MarkerRectangles(result.Page, new RgbColor(255, 242, 204));
            TestAssert.Equal(count, backgrounds.Length);
            for (int index = 1; index < count; index++)
                TestAssert.True(Math.Abs((backgrounds[index].Y - baselines[index]) - (backgrounds[0].Y - baselines[0])) < .02d,
                    "Each background follows the same placement as its text.");
            if (key)
            {
                var keys = PptxChartLabelDefaultsTests.MarkerRectangles(result.Page, new RgbColor(255, 0, 254)).Take(count).ToArray();
                TestAssert.Equal(count, keys.Length);
                for (int index = 1; index < count; index++)
                    TestAssert.True(Math.Abs((keys[index].Y - baselines[index]) - (keys[0].Y - baselines[0])) < .02d,
                        "Each key follows its complete label unit.");
            }
            if (rich)
                for (int index = 1; index <= count; index++)
                {
                    var annotation = result.Page.Annotations.Single(item => item.Uri == "https://example.invalid/label/" + index);
                    TestAssert.True(annotation.Y <= groups[index - 1].Y + .02d &&
                        annotation.Y + annotation.Height >= groups[index - 1].Y + groups[index - 1].Height - .02d,
                        "Hyperlinks follow the final painted label's vertical frame.");
                }
            TestAssert.True(result.Diagnostics.All(item => item.Id != "PPTX_CHART_LABEL_OVERLAP"), "These small label arrangements fit their admitted rows.");
        }
    }

    public static void PptxPolarLabelAdmissionRetainsAuthoredAndUnfitArrangements()
    {
        const string edge = "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>" +
            "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".2\"/><c:y val=\".2\"/>" +
            "<c:w val=\".2\"/><c:h val=\".1\"/></c:manualLayout></c:layout>";
        foreach (bool authored in new[] { false, true })
        {
            string points = authored ? string.Concat(Enumerable.Range(0, 6).Select(i =>
                $"<c:dLbl><c:idx val=\"{i}\"/>" + edge + AdmissionFlags + "<c:showLegendKey val=\"0\"/></c:dLbl>")) : "";
            string labels = "<c:dLbls>" + points + AdmissionFlags + "<c:showLegendKey val=\"0\"/></c:dLbls>";
            var result = Render(576, 12, 6, false, kind: "pie", height: authored ? 360 : 36,
                categoryLegend: true, plotLayout: AdmissionPlot, dataLabels: labels,
                seriesOptions: "<c:tx><c:v>Series 1</c:v></c:tx><c:explosion val=\"20\"/>");
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(result.Page, false,
                fontResourceFilter: name => name.StartsWith("CP", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(6, shows.Length);
            TestAssert.Equal(1, result.Diagnostics.Count(item => item.Id == "PPTX_CHART_LABEL_OVERLAP"));
            if (authored) TestAssert.True(shows.All(show => Math.Abs(show.Y - shows[0].Y) < .001d),
                "Conflicting authored label coordinates remain fixed rather than being silently rewritten.");
        }
    }

    private static PdfRectangle UnionLabelInk(PdfRectangle a, PdfRectangle b)
    {
        double x = Math.Min(a.X, b.X), y = Math.Min(a.Y, b.Y);
        return new(x, y, Math.Max(a.X + a.Width, b.X + b.Width) - x, Math.Max(a.Y + a.Height, b.Y + b.Height) - y);
    }

    private static bool LabelInkOverlaps(PdfRectangle a, PdfRectangle b) =>
        Math.Min(a.X + a.Width, b.X + b.Width) > Math.Max(a.X, b.X) + .02d &&
        Math.Min(a.Y + a.Height, b.Y + b.Height) > Math.Max(a.Y, b.Y) + .02d;
}
