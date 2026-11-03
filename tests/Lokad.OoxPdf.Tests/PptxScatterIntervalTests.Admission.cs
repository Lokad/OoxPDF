using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterCaptionAdmissionPreservesFrameAndNumericCenters()
    {
        foreach (string id in new[] { "auto-x-width-360-long-format", "auto-x-inner-0p95" })
        foreach (bool reversed in new[] { false, true })
        {
            AutoCase sample = AutoCases().Single(c => c.Id == id);
            XDocument xml = XDocument.Parse(AutoChart(sample));
            if (reversed) xml.Descendants(CurveC + "valAx").First().Element(CurveC + "scaling")!
                .Element(CurveC + "orientation")!.SetAttributeValue("val", "maxMin");
            PdfPage page = RenderAuto(sample, xml.ToString());
            double[] axis = AdmissionAxis(page);
            double width = axis[2] - axis[0];
            double sourceWidth = sample.FrameWidth * sample.WidthFactor;
            if (id.Contains("long-format", StringComparison.Ordinal)) TestAssert.True(Math.Abs(width - sourceWidth) < .002d,
                "A fitting source width must translate rather than shrink.");
            else TestAssert.True(width < sourceWidth - 1d, "An oversized source plot must shrink for its caption slots.");
            int count = (int)Math.Round((sample.ExpectedMaximum - sample.ExpectedMinimum) / sample.ExpectedUnit) + 1;
            var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Take(count).ToArray();
            TestAssert.Equal(count, labels.Length);
            foreach (var label in labels)
            {
                double advance = label.Font.MeasureTextPoints(label.Text, sample.FontSize, kerningEnabled: false);
                TestAssert.True(label.X >= 72d && label.X + advance <= 72d + sample.FrameWidth + .01d,
                    "Every endpoint caption must stay inside the chart frame.");
                double value = double.Parse(label.Text, CultureInfo.InvariantCulture);
                double ratio = (value - sample.ExpectedMinimum) / (sample.ExpectedMaximum - sample.ExpectedMinimum);
                if (reversed) ratio = 1d - ratio;
                TestAssert.True(Math.Abs(label.X + advance / 2d - (axis[0] + width * ratio)) < .05d,
                    "A wide numeric caption must retain its tick center.");
            }
        }
    }

    public static void PptxScatterCaptionAdmissionKeepsEndpointAndSeriesOwnership()
    {
        // Independent Office controls keep X99.132..603.230 when the unpainted
        //103 endpoint becomes a painted107 tick or the series line is hidden;
        // high-side labels instead use X80.767..584.870 at the same plot width.
        AutoCase sample = AutoCases().Single(c => c.Id == "auto-x-inner-0p95");
        PdfPage baseline = RenderAuto(sample, AutoChart(sample));
        double[] source = AdmissionAxis(baseline);
        TestAssert.True(source[2] - source[0] < sample.FrameWidth * sample.WidthFactor - 1d,
            "The base source needs caption admission.");
        foreach (string variant in new[] { "aligned", "high", "hidden" })
        {
            XDocument xml = XDocument.Parse(AutoChart(sample));
            XElement yAxis = xml.Descendants(CurveC + "valAx").Last();
            if (variant == "aligned") yAxis.Element(CurveC + "scaling")!.Element(CurveC + "max")!.SetAttributeValue("val", "107");
            if (variant == "high") yAxis.Add(new XElement(CurveC + "tickLblPos", new XAttribute("val", "high")));
            if (variant == "hidden")
            {
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                xml.Descendants(CurveC + "ser").Single().AddFirst(new XElement(CurveC + "spPr", new XElement(a + "ln", new XElement(a + "noFill"))));
            }
            PdfPage page = RenderAuto(sample, xml.ToString());
            double[] axis = AdmissionAxis(page);
            TestAssert.True(Math.Abs((axis[2] - axis[0]) - (source[2] - source[0])) < .002d,
                "Endpoint visibility, label side and series paint retain the same admitted width.");
            if (variant == "high") TestAssert.True(axis[0] < source[0] - 1d && axis[2] < source[2] - 1d,
                "High-side captions move the reserve to the other frame edge.");
            else TestAssert.True(axis.Zip(source).All(pair => Math.Abs(pair.First - pair.Second) < .002d),
                "The unpainted endpoint participates in measurement; series visibility does not own axis placement.");
            if (variant != "aligned") TestAssert.Equal(string.Join(",", AutoLabels(baseline)), string.Join(",", AutoLabels(page)));
        }
    }

    private static double[] AdmissionAxis(PdfPage page) => PptxColumnIntervalTests.Lines(page).Single(line =>
        line.Color == default && Math.Abs(line.V[1] - line.V[3]) < .001d && Math.Abs(line.V[0] - line.V[2]) > 100d).V;
}
