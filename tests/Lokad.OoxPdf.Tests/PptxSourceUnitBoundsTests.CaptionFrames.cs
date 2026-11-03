using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxSourceUnitBoundsTests
{
    public static void PptxObliqueCategoryCaptionsAdmitTheirHorizontalFrames()
    {
        foreach (int angle in new[] { -45, 45 })
        foreach (int unit in new[] { 1, 2, 5 })
        {
            PdfPage page = RenderCaptionFrames(angle, unit);
            CheckCaptionInk(page, angle);
            double maximum = unit == 2 ? 6 : 5;
            double[] labels = Labels(page);
            TestAssert.True(labels[0] == 0 && labels[^1] == maximum,
                "Admitting category frames preserves the numeric source scale.");
        }
    }

    public static void PptxObliqueCategoryFramesShareDirectionAndEndpointPlacement()
    {
        foreach (int angle in new[] { -45, 45 })
        foreach (bool reversed in new[] { false, true })
        foreach (bool highSide in new[] { false, true })
        foreach (string kind in new[] { "bar", "line" })
            CheckCaptionInk(RenderCaptionFrames(angle, 2, kind, reversed, highSide), angle);
    }

    public static void PptxCategoryFrameAdmissionPreservesManualAndVerticalOwners()
    {
        foreach (int angle in new[] { -45, 45 })
        {
            var lines = PptxColumnIntervalTests.Lines(RenderCaptionFrames(angle, 2, manual: true));
            double[] axis = lines.Single(line => line.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(line.V[0] - line.V[2]) < .001 && Math.Abs(line.V[1] - line.V[3]) > 100).V;
            TestAssert.True(Math.Abs(axis[0] - (72 + .1 * 576)) < .002,
                "A manual inner plot keeps its authored horizontal origin.");
        }
        double[] Axis(PdfPage page) => PptxColumnIntervalTests.Lines(page).Single(line => line.Color == new RgbColor(255, 0, 0) &&
            Math.Abs(line.V[0] - line.V[2]) < .001 && Math.Abs(line.V[1] - line.V[3]) > 100).V;
        TestAssert.True(Math.Abs(Axis(RenderCaptionFrames(-90, 2))[0] - Axis(RenderCaptionFrames(-90, 2, hidden: true))[0]) < .002,
            "Vertical captions add no horizontal advance reserve, matching the native control.");
    }

    private static void CheckCaptionInk(PdfPage page, int angle)
    {
        var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal(19, captions.Length);
        double radians = -angle * Math.PI / 180, cos = Math.Cos(radians), sin = Math.Sin(radians);
        foreach (var show in captions)
        {
            double cursor = 0, left = double.PositiveInfinity, right = double.NegativeInfinity;
            var font = show.Font.Font;
            foreach (Rune rune in show.Text.EnumerateRunes())
            {
                ushort glyph = font.MapCodePoint(rune.Value);
                TestAssert.True(font.TryGetGlyphBounds(glyph, out var bounds), "Expected pinned category glyph geometry.");
                double scale = show.Size / font.UnitsPerEm;
                foreach (double x in new[] { bounds.XMin * scale, bounds.XMax * scale })
                foreach (double y in new[] { bounds.YMin * scale, bounds.YMax * scale })
                {
                    double actual = show.X + (cursor + x) * cos - y * sin;
                    left = Math.Min(left, actual);
                    right = Math.Max(right, actual);
                }
                cursor += font.GetAdvanceWidth(glyph) * scale;
            }
            TestAssert.True(left >= 72 && right <= 648,
                $"Full category ink survives the chart clip: {show.Text} {left:F3}..{right:F3}.");
        }
    }

    private static PdfPage RenderCaptionFrames(int angle, int unit, string kind = "bar", bool reversed = false,
        bool highSide = false, bool manual = false, bool hidden = false)
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var sample = new Sample(Enumerable.Repeat(4d, 19).ToArray(), unit, null, 0, unit == 2 ? 6 : 5);
        XElement xml = XElement.Parse(Chart(sample, kind, false, false));
        XElement plot = xml.Descendants(c + "plotArea").Single();
        XElement ordinal = plot.Element(c + "catAx")!;
        ordinal.Element(c + "axId")!.AddAfterSelf(new XElement(c + "scaling",
            new XElement(c + "orientation", new XAttribute("val", reversed ? "maxMin" : "minMax"))));
        ordinal.Add(new XElement(c + "tickLblPos", new XAttribute("val", hidden ? "none" : highSide ? "high" : "nextTo")));
        ordinal.Add(new XElement(c + "txPr", new XElement(a + "bodyPr", new XAttribute("rot", angle * 60000)),
            new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", 1200))))));
        XElement numeric = plot.Element(c + "valAx")!;
        numeric.Element(c + "numFmt")!.SetAttributeValue("formatCode", "0");
        if (kind == "line") numeric.Add(new XElement(c + "crossBetween", new XAttribute("val", "midCat")));
        if (manual) plot.AddFirst(XElement.Parse("""
            <c:layout xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val=".1"/><c:y val=".1"/><c:w val=".8"/><c:h val=".5"/></c:manualLayout></c:layout>
            """));
        const string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="7315200" cy="4572000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(xml.ToString(SaveOptions.DisableFormatting),
            new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
