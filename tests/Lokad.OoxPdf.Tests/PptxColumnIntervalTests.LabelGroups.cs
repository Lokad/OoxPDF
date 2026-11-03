using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    public static void PptxCenteredHorizontalLabelFramesFollowPaintedBars()
    {
        foreach (int count in new[] { 1, 2 })
        foreach (string grouping in new[] { "clustered", "stacked" })
        {
            PdfPage page = LabelGroupPage(seriesCount: count, grouping: grouping);
            var bars = LabelGroupRectangles(page, "0.2 0.4 0.8 rg").Where(r => r[2] > 20d && r[3] > 20d).ToArray();
            var frames = LabelGroupRectangles(page, "1 1 0 rg");
            TestAssert.Equal(2 * count, bars.Length);
            TestAssert.Equal(bars.Length, frames.Length);
            // Painting orders stacked series before categories, while clustered
            // bars and labels order categories first. Match observable centers.
            foreach (double[] box in frames)
                TestAssert.True(bars.Any(bar => Math.Abs(box[0] + box[2] / 2d - bar[0] - bar[2] / 2d) < .002d &&
                    Math.Abs(box[1] + box[3] / 2d - bar[1] - bar[3] / 2d) < .002d),
                    "The full label frame follows an actual painted rectangle, including grouped spacing.");
        }
    }

    public static void PptxCenteredHorizontalLabelWidthFollowsContentInsteadOfPlotWidth()
    {
        foreach (bool key in new[] { false, true })
        {
            double Width(string text, double plotWidth) => LabelGroupRectangles(LabelGroupPage(key: key, name: text, plotWidth: plotWidth), "1 1 0 rg")[0][2];
            double narrow = Width("Actual", .45d), wide = Width("Actual", .8d);
            NearHorizontal(narrow, wide);
            TestAssert.True(Width("A much longer source series name", .8d) > wide + 20d,
                "Content advance determines occupied width; changing the plot rectangle does not stretch the label.");
        }
    }

    public static void PptxCenteredHorizontalLabelKeysAndTextFitTheirMeasuredFrames()
    {
        foreach (int fontSize in new[] { 12, 18 })
        foreach (bool key in new[] { false, true })
        {
            PdfPage page = LabelGroupPage(fontSize: fontSize, key: key);
            var frames = LabelGroupRectangles(page, "1 1 0 rg");
            var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(2, labels.Length);
            for (int i = 0; i < labels.Length; i++)
            {
                double advance = labels[i].Font.MeasureTextPoints(labels[i].Text, labels[i].Size);
                TestAssert.True(labels[i].X > frames[i][0] && labels[i].X + advance < frames[i][0] + frames[i][2] &&
                    labels[i].Y > frames[i][1] && labels[i].Y < frames[i][1] + frames[i][3],
                    "Measured text and its baseline fit inside the occupied frame.");
            }
            if (key)
            {
                var keys = LabelGroupRectangles(page, "0.2 0.4 0.8 rg").Where(r => r[2] < 20d && r[3] < 20d).ToArray();
                TestAssert.Equal(2, keys.Length);
                TestAssert.True(keys.Zip(labels).All(p => p.First[0] + p.First[2] < p.Second.X), "The key has its own space before the text.");
            }
        }
    }

    public static void PptxCenteredHorizontalLabelManualAndCustomFallbacksKeepContent()
    {
        PdfPage Center(string position) => LabelGroupPage(position: position, extra: labels => labels.Add(XElement.Parse(
            "<c:layout xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".2\"/><c:y val=\".3\"/><c:w val=\".4\"/><c:h val=\".15\"/></c:manualLayout></c:layout>")));
        TestAssert.Equal(Center("ctr").Content, Center("outEnd").Content);
        PdfPage custom = LabelGroupPage(extra: labels =>
        {
            for (int index = 0; index < 2; index++) labels.Add(new XElement(CrossingC + "dLbl",
                new XElement(CrossingC + "idx", new XAttribute("val", index)), XElement.Parse(
                "<c:tx xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"1200\"/><a:t>Custom </a:t></a:r><a:r><a:rPr sz=\"1800\" b=\"1\"/><a:t>label</a:t></a:r></a:p></c:rich></c:tx>")));
        });
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(custom, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).Select(p => p.Text));
        TestAssert.Equal("Custom labelCustom label", text);
    }

    private static PdfPage LabelGroupPage(int seriesCount = 1, string grouping = "clustered", bool key = true,
        string name = "Actual", double plotWidth = .8d, int fontSize = 18, string position = "ctr", Action<XElement>? extra = null)
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        return Render(0, 25, maximum: 200, grouping: grouping, horizontal: true, editChart: xml =>
        {
            var bar = xml.Descendants(CrossingC + "barChart").Single();
            var series = bar.Element(CrossingC + "ser")!;
            series.Add(new XElement(CrossingC + "tx", new XElement(CrossingC + "v", name)),
                new XElement(CrossingC + "spPr", new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "3366CC")))));
            if (seriesCount == 2)
            {
                var second = new XElement(series);
                second.Element(CrossingC + "idx")!.SetAttributeValue("val", "1");
                second.Element(CrossingC + "order")!.SetAttributeValue("val", "1");
                bar.Add(second);
            }
            var area = xml.Descendants(CrossingC + "plotArea").Single();
            area.Add(new XElement(CrossingC + "layout", new XElement(CrossingC + "manualLayout",
                new XElement(CrossingC + "layoutTarget", new XAttribute("val", "inner")),
                new XElement(CrossingC + "xMode", new XAttribute("val", "factor")),new XElement(CrossingC + "yMode",new XAttribute("val","factor")),
                new XElement(CrossingC + "wMode",new XAttribute("val","factor")),new XElement(CrossingC + "hMode",new XAttribute("val","factor")),
                new XElement(CrossingC + "x",new XAttribute("val",".1")),new XElement(CrossingC + "y",new XAttribute("val",".1")),
                new XElement(CrossingC + "w",new XAttribute("val",plotWidth.ToString("R",CultureInfo.InvariantCulture))),new XElement(CrossingC + "h",new XAttribute("val",".8")))));
            var labels = new XElement(CrossingC + "dLbls",new XElement(CrossingC + "dLblPos",new XAttribute("val",position)),
                new XElement(CrossingC + "showVal",new XAttribute("val","1")),new XElement(CrossingC + "showLegendKey",new XAttribute("val",key?"1":"0")),
                new XElement(CrossingC + "spPr",new XElement(a+"solidFill",new XElement(a+"srgbClr",new XAttribute("val","FFFF00")))),
                new XElement(CrossingC+"txPr",new XElement(a+"bodyPr"),new XElement(a+"lstStyle"),new XElement(a+"p",new XElement(a+"pPr",new XElement(a+"defRPr",new XAttribute("sz",fontSize*100))))));
            extra?.Invoke(labels);bar.Add(labels);
        });
    }

    private static double[][] LabelGroupRectangles(PdfPage page, string selectedColor)
    {
        var result = new List<double[]>();var colors = new Stack<string>();string color = "";
        var points = new List<double[]>();var pending = new List<double[]>();
        foreach (string line in page.Content.Split('\n').Select(s => s.Trim()))
        {
            if (line == "q") colors.Push(color);else if (line == "Q") color=colors.Pop();
            else if (line.EndsWith(" rg",StringComparison.Ordinal) || line.EndsWith(" g",StringComparison.Ordinal)) color=line;
            else
            {
                var match=Regex.Match(line,@"^([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f$");
                if (match.Success && color == selectedColor) result.Add(match.Groups.Cast<Group>().Skip(1).Select(g=>double.Parse(g.Value,CultureInfo.InvariantCulture)).ToArray());
                var point=Regex.Match(line,@"^([-\d.]+) ([-\d.]+) [ml]$");
                if (point.Success) points.Add([double.Parse(point.Groups[1].Value,CultureInfo.InvariantCulture),double.Parse(point.Groups[2].Value,CultureInfo.InvariantCulture)]);
                else if (line == "h")
                {
                    if (points.Count == 4)
                    {
                        double left=points.Min(p=>p[0]),bottom=points.Min(p=>p[1]);
                        pending.Add([left,bottom,points.Max(p=>p[0])-left,points.Max(p=>p[1])-bottom]);
                    }
                    points.Clear();
                }
                else if (line is "f" or "f*")
                {
                    if (color == selectedColor) result.AddRange(pending);
                    pending.Clear();points.Clear();
                }
                else if (line is "n" or "S") {pending.Clear();points.Clear();}
            }
        }
        return result.ToArray();
    }
}
