using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxSourceUnitBoundsTests
{
    public static void PptxObliqueCategoryCaptionsKeepSourceScaleWithinTheirChartClip()
    {
        foreach (int angle in new[] { -45, 45 })
        foreach (int unit in new[] { 1, 2, 5 })
        {
            PdfPage page = RenderCaptionFrames(angle, unit);
            CheckChartCaptionClip(page);
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
            CheckChartCaptionClip(RenderCaptionFrames(angle, 2, kind, reversed, highSide));
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

    internal static void CheckCaptionInk(PdfPage page, int angle, int expectedCount = 19,
        double frameLeft = 72d, double frameRight = 648d, double? frameBottom = null, double? frameTop = null)
    {
        var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal(expectedCount, captions.Length);
        double radians = -angle * Math.PI / 180, cos = Math.Cos(radians), sin = Math.Sin(radians);
        foreach (var show in captions)
        {
            double cursor = 0, left = double.PositiveInfinity, right = double.NegativeInfinity;
            double bottom = double.PositiveInfinity, top = double.NegativeInfinity;
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
                    double vertical = show.Y + (cursor + x) * sin + y * cos;
                    bottom = Math.Min(bottom, vertical);
                    top = Math.Max(top, vertical);
                }
                cursor += font.GetAdvanceWidth(glyph) * scale;
            }
            TestAssert.True(left >= frameLeft && right <= frameRight &&
                (frameBottom is null || bottom >= frameBottom) && (frameTop is null || top <= frameTop),
                $"Full category ink survives the chart clip: {show.Text} {left:F3}..{right:F3}.");
        }
    }

    private static PdfPage RenderCaptionFrames(int angle, int unit, string kind = "bar", bool reversed = false,
        bool highSide = false, bool manual = false, bool hidden = false, int height = 360,
        int categoryFontSize = 1200, string? overflow = null, string? categoryText = null, int valueFontSize = 1000,
        int? builtInStyle = null, int? categoryRoleFontSize = null, int? chartFontSize = null,
        int? valueRoleFontSize = null, bool inheritAxisFontSizes = false, double manualHeight = .5d, int? labelSkip = null,
        bool? onTickMarks = null, int? interiorCaptionIndex = null, double manualY = .1d,
        bool valueReversed = false, int width = 576, string? firstCaptionText = null, string? wrap = null)
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var sample = new Sample(Enumerable.Repeat(4d, 19).ToArray(), unit, null, 0, unit == 2 ? 6 : 5);
        XElement xml = XElement.Parse(Chart(sample, kind, false, false));
        XElement plot = xml.Descendants(c + "plotArea").Single();
        if (categoryText is not null)
            foreach (XElement point in plot.Descendants(c + "cat").Descendants(c + "pt"))
                point.Element(c + "v")!.Value = interiorCaptionIndex is null || (int)point.Attribute("idx")! == interiorCaptionIndex
                    ? categoryText : "I";
        if (firstCaptionText is not null)
            plot.Descendants(c + "cat").Descendants(c + "pt").First().Element(c + "v")!.Value = firstCaptionText;
        XElement ordinal = plot.Element(c + "catAx")!;
        if (labelSkip is { } explicitSkip)
        {
            ordinal.Elements(c + "tickLblSkip").Remove();
            ordinal.Add(new XElement(c + "tickLblSkip", new XAttribute("val", explicitSkip)));
        }
        ordinal.Element(c + "axId")!.AddAfterSelf(new XElement(c + "scaling",
            new XElement(c + "orientation", new XAttribute("val", reversed ? "maxMin" : "minMax"))));
        ordinal.Add(new XElement(c + "tickLblPos", new XAttribute("val", hidden ? "none" : valueReversed
            ? highSide ? "low" : "high" : highSide ? "high" : "nextTo")));
        ordinal.Add(new XElement(c + "txPr", new XElement(a + "bodyPr", new XAttribute("rot", angle * 60000),
            overflow is null ? null : new XAttribute("vertOverflow", overflow), wrap is null ? null : new XAttribute("wrap", wrap)),
            new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", categoryFontSize))))));
        XElement numeric = plot.Element(c + "valAx")!;
        if (valueReversed)
        {
            numeric.Element(c + "scaling")!.Elements(c + "orientation").Remove();
            numeric.Element(c + "scaling")!.AddFirst(new XElement(c + "orientation", new XAttribute("val", "maxMin")));
        }
        numeric.Descendants(a + "defRPr").Single().SetAttributeValue("sz", valueFontSize);
        numeric.Element(c + "numFmt")!.SetAttributeValue("formatCode", "0");
        if (kind == "line" || onTickMarks is not null)
            numeric.Add(new XElement(c + "crossBetween", new XAttribute("val", (onTickMarks ?? true) ? "midCat" : "between")));
        if (manual) plot.AddFirst(XElement.Parse("""
            <c:layout xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val=".1"/><c:y val=".1"/><c:w val=".8"/><c:h val=".5"/></c:manualLayout></c:layout>
            """));
        if (manual)
        {
            var source = plot.Descendants(c + "manualLayout").Single();
            source.Element(c + "h")!.SetAttributeValue("val", manualHeight.ToString(CultureInfo.InvariantCulture));
            source.Element(c + "y")!.SetAttributeValue("val", manualY.ToString(CultureInfo.InvariantCulture));
        }
        string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="7315200" cy="4572000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
            """;
        slide = slide.Replace("cy=\"4572000\"", $"cy=\"{height * 12700}\"", StringComparison.Ordinal);
        slide = slide.Replace("cx=\"7315200\"", $"cx=\"{width * 12700}\"", StringComparison.Ordinal);
        if (builtInStyle is { } styleId)
            xml.AddFirst(new XElement(c + "style", new XAttribute("val", styleId)));
        if (chartFontSize is { } globalSize)
            xml.Add(new XElement(c + "txPr", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"),
                new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", globalSize))))));
        if (inheritAxisFontSizes)
            foreach (XElement axis in plot.Elements().Where(element => element.Name == c + "catAx" || element.Name == c + "valAx"))
                foreach (XElement properties in axis.Descendants(a + "defRPr")) properties.Attribute("sz")?.Remove();
        var parts = new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) };
        if (builtInStyle is not null)
        {
            // The measured frame needs a named theme face. The resolver pins its
            // program on either host; missing theme metadata exercises a fallback.
            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement presentationRelationships = XElement.Parse(PptxTests.PresentationRelationship());
            presentationRelationships.Add(new XElement(relationships + "Relationship",
                new XAttribute("Id", "rTheme"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                new XAttribute("Target", "theme/theme1.xml")));
            parts["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(presentationRelationships.ToString(SaveOptions.DisableFormatting));
            parts["ppt/theme/theme1.xml"] = TestFixtures.Utf8(new XElement(a + "theme",
                new XElement(a + "themeElements", new XElement(a + "fontScheme",
                    new XElement(a + "majorFont", new XElement(a + "latin", new XAttribute("typeface", "DejaVu Sans"))),
                    new XElement(a + "minorFont", new XElement(a + "latin", new XAttribute("typeface", "DejaVu Sans")))))).ToString(SaveOptions.DisableFormatting));
        }
        if (categoryRoleFontSize is not null || valueRoleFontSize is not null)
        {
            parts["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/></Relationships>
                """);
            XNamespace cs = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
            parts["ppt/charts/style1.xml"] = TestFixtures.Utf8(new XElement(cs + "chartStyle",
                new XAttribute("id", "201"),
                categoryRoleFontSize is { } roleSize ? new XElement(cs + "categoryAxis",
                    new XElement(cs + "defRPr", new XAttribute("sz", roleSize))) : null,
                valueRoleFontSize is { } valueSize ? new XElement(cs + "valueAxis",
                    new XElement(cs + "defRPr", new XAttribute("sz", valueSize))) : null).ToString(SaveOptions.DisableFormatting));
        }
        using var stream = PptxTests.CreateSingleChartPackageStream(xml.ToString(SaveOptions.DisableFormatting), parts);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
