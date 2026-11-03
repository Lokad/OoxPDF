using System.Globalization;
using System.Buffers.Binary;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxRightLegendAdmissionTests
{
    public static void PptxAreaSideLegendPaintInsetPreservesTextGapAndFrameOwnership()
    {
        foreach (ushort ascender in new ushort[] { 900, 1300 })
        foreach (int size in new[] { 12, 24, 1 })
        foreach (bool frame in new[] { false, true })
        {
            PdfPage page = Render(true, "high", false, legendFontSize: size,
                resolver: new LegendLineBoxResolver(ascender), legendNames: ["Revenue", "gyp"], legendFrame: frame);
            var keys = Keys(page);
            TestAssert.Equal(2, keys.Length);
            foreach (var (key, name) in keys.Zip(new[] { "Revenue", "gyp" }))
            {
                double gap = TextOrigin(page, name).X - key.X - key.Size;
                TestAssert.True(gap >= -.002d && Math.Abs(gap - Math.Max(0d, key.Size / 2d - .25d)) < .002d,
                    "The painted key uses its measured slot and fixed native inset without overlapping its label.");
            }
            TestAssert.Equal(1, DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Count(s => s.Text == "gyp"));
        }
    }

    public static void PptxAreaSideLegendPaintedKeysAndRowsUseThePreparedLineBox()
    {
        PdfPage shortFont = Render(true, "high", false, legendFontSize: 18,
            resolver: new LegendLineBoxResolver(900), legendNames: ["Revenue", "gyp"]);
        PdfPage tallFont = Render(true, "high", false, legendFontSize: 18,
            resolver: new LegendLineBoxResolver(1300), legendNames: ["Revenue", "gyp"]);
        var shortKeys = Keys(shortFont);
        var tallKeys = Keys(tallFont);
        TestAssert.Equal(2, shortKeys.Length);
        TestAssert.Equal(2, tallKeys.Length);
        TestAssert.True(tallKeys[0].Size - shortKeys[0].Size is > 3d and < 4d,
            "Identical glyph advances at the same font size still require keys from the selected font line box.");
        double shortPitch = shortKeys[0].CenterY - shortKeys[1].CenterY;
        double tallPitch = tallKeys[0].CenterY - tallKeys[1].CenterY;
        TestAssert.True(tallPitch - shortPitch is > 7d and < 8d,
            "Prepared rows follow normal line metrics instead of a scalar font-size regime.");
        TestAssert.True(Math.Abs(TextOrigin(shortFont, "Revenue").X - TextOrigin(tallFont, "Revenue").X) < .002d,
            "Measured key/row geometry preserves the anchored text advance.");
        foreach (var page in new[] { shortFont, tallFont })
        {
            AssertText(page, 3, "W");
            TestAssert.Equal(1, DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Count(s => s.Text == "gyp"));
        }
    }

    public static void PptxAreaSideLegendRowsKeepOneFontBaselineAcrossDifferentInk()
    {
        PdfPage page = Render(true, "high", false, legendFontSize: 18, legendNames: ["Revenue", "gyp"]);
        var keys = Keys(page);
        TestAssert.Equal(2, keys.Length);
        var first = TextOrigin(page, "Revenue");
        var last = TextOrigin(page, "gyp");
        TestAssert.True(Math.Abs(first.X - last.X) < .002d &&
            Math.Abs(first.Y - last.Y - (keys[0].CenterY - keys[1].CenterY)) < .002d,
            "Independent Office mixed-ink rows retain the same font baseline offset; descenders do not recenter each entry.");
        AssertText(page, 3, "W");
    }

    private static (double X, double Size, double CenterY)[] Keys(PdfPage page) =>
        Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re f\*?")
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(r => r[0] > 600d && r[2] is > 0d and < 20d && Math.Abs(r[2] - r[3]) < .002d)
            .Select(r => (X: r[0], Size: r[2], CenterY: r[1] + r[3] / 2d)).OrderByDescending(r => r.CenterY).ToArray();

    public static void PptxAreaRightLegendUsesFontLineBoxWhenAdvancesStayFixed()
    {
        PdfPage shortFont = Render(true, "high", false, legendFontSize: 18, resolver: new LegendLineBoxResolver(900));
        PdfPage tallFont = Render(true, "high", false, legendFontSize: 18, resolver: new LegendLineBoxResolver(1300));
        var shortSpan = Span(shortFont);
        var tallSpan = Span(tallFont);
        TestAssert.True(Math.Abs(shortSpan.Left - tallSpan.Left) < .002d &&
            shortSpan.Right - tallSpan.Right is > 5d and < 8d,
            "Same-size fonts with identical glyph advances need different side-legend allocations when their normal line boxes differ.");
        TestAssert.True(Math.Abs(TextOrigin(shortFont, "Revenue").X - TextOrigin(tallFont, "Revenue").X) < .002d,
            "Allocating the legend's measured frame keeps its independently anchored text in place.");
        AssertText(shortFont, 3, "W");
        AssertText(tallFont, 3, "W");
    }

    public static void PptxAreaRightLegendLineMetricsKeepLineAndAuthoredFallbacks()
    {
        foreach (string? fallback in new string?[] { null, "manual", "rotated", "unpaired", "log" })
        {
            // The null case selects a line chart: the new area allocation must
            // not silently replace its stroke-key reserve.
            bool area = fallback is not null;
            PdfPage shortFont = Render(area, "high", false, fallback: fallback, legendFontSize: 18, resolver: new LegendLineBoxResolver(900));
            PdfPage tallFont = Render(area, "high", false, fallback: fallback, legendFontSize: 18, resolver: new LegendLineBoxResolver(1300));
            TestAssert.Equal(Span(shortFont), Span(tallFont));
            TestAssert.True(Keys(shortFont).SequenceEqual(Keys(tallFont)), "Fallback key geometry keeps its existing owner.");
            AssertText(shortFont, 3, "W");
            AssertText(tallFont, 3, "W");
        }
    }

    public static void PptxRightLegendAdmissionMovesNumericBandsWithoutMovingLegend()
    {
        foreach (bool area in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        {
            PdfPage low = Render(area, "low", reversed);
            PdfPage high = Render(area, "high", reversed);
            var left = Span(reversed ? high : low);
            var right = Span(reversed ? low : high);
            TestAssert.True(left.Left > right.Left + 1d && left.Right > right.Right + 1d,
                "Numeric width belongs beside the painted labels; moving it shifts both plot edges.");
            TestAssert.True(Math.Abs(left.Right - left.Left - (right.Right - right.Left)) < .002d,
                "Equal caption endpoints and the same numeric band preserve the inner width.");
            var lowLegend = TextOrigin(low, "Revenue");
            var highLegend = TextOrigin(high, "Revenue");
            TestAssert.True(Math.Abs(lowLegend.X - highLegend.X) < .002d && Math.Abs(lowLegend.Y - highLegend.Y) < .002d,
                "The independent legend container retains its position when numeric labels change side.");
            AssertText(low, 3, "W");
            AssertText(high, 3, "W");
        }
    }

    public static void PptxRightLegendAdmissionUsesCaptionInkAndSourceSlots()
    {
        foreach (int count in new[] { 1, 3, 8 })
        {
            var between = Span(Render(false, "high", false, count: count, between: true));
            TestAssert.True(Math.Abs(between.Left - 83d) < .002d,
                "Between-slot captions fit inside their source slots without an endpoint half-caption band.");
        }
        PdfPage small = Render(false, "high", false, count: 1, fontSize: 12);
        PdfPage large = Render(false, "high", false, count: 1, fontSize: 18);
        PdfPage wide = Render(false, "high", false, count: 1, caption: "WideCaption");
        foreach (var (page, prefix) in new[] { (small, "W"), (large, "W"), (wide, "WideCaption") })
        {
            var (left, _) = Span(page);
            string text = prefix + "1";
            TestAssert.True(Math.Abs(TextOrigin(page, text).X - 83d) < .002d,
                "The first tick-centered caption starts at the existing frame inset, using its actual font advance.");
            TestAssert.True(left > 83d, "Endpoint captions reserve their actual half advance.");
            AssertText(page, 1, prefix);
        }
        TestAssert.True(Span(large).Left > Span(small).Left + 1d && Span(wide).Left > Span(large).Left + 1d,
            "Font size and literal caption width drive admission, rather than a fixed translation.");
        TestAssert.True(Span(wide).Right < Span(large).Right - 1d,
            "A wider endpoint caption competes with numeric labels on the right too.");
    }

    public static void PptxRightLegendAdmissionKeepsAuthoredAndRotatedFallbacks()
    {
        foreach (bool area in new[] { false, true })
        foreach (string fallback in new[] { "manual", "rotated", "unpaired", "log" })
        {
            PdfPage low = Render(area, "low", false, fallback: fallback);
            PdfPage high = Render(area, "high", false, fallback: fallback);
            TestAssert.Equal(Span(low), Span(high));
            TestAssert.Equal(TextOrigin(low, "Revenue"), TextOrigin(high, "Revenue"));
            AssertText(low, 3, "W");
            AssertText(high, 3, "W");
        }
    }

    private static void AssertText(PdfPage page, int count, string prefix)
    {
        string[] text = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false, includeTextTransforms: true).Select(s => s.Text).ToArray();
        TestAssert.Equal(1, text.Count(s => s == "Revenue"));
        for (int i = 1; i <= count; i++) TestAssert.Equal(1, text.Count(s => s == prefix + i));
    }

    private static (double X, double Y) TextOrigin(PdfPage page, string text)
    {
        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false, includeTextTransforms: true).ToArray();
        var matches = shows.Where(s => s.Text == text).ToArray();
        TestAssert.True(matches.Length == 1, $"Expected one {text}; got {matches.Length}. Shows: {string.Join("|", shows.Select(s => s.Text))}");
        var show = matches[0];
        return (show.X, show.Y);
    }

    private static (double Left, double Right) Span(PdfPage page)
    {
        var lines = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) m ([-\d.]+) ([-\d.]+) l S")
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(p => Math.Abs(p[1] - p[3]) < .001d && Math.Abs(p[0] - p[2]) > 200d).ToArray();
        TestAssert.True(lines.Length > 0, "Inspect the painted horizontal plot span.");
        return lines.GroupBy(p => (Left: Math.Min(p[0], p[2]), Right: Math.Max(p[0], p[2])))
            .OrderByDescending(g => g.Count()).First().Key;
    }

    private static PdfPage Render(bool area, string side, bool reversed, int count = 3, bool between = false,
        int fontSize = 18, string caption = "W", string? fallback = null,
        int? legendFontSize = null, IFontResolver? resolver = null, IReadOnlyList<string>? legendNames = null, bool legendFrame = false)
    {
        string categories = string.Concat(Enumerable.Range(0, count).Select(i => $"<c:pt idx=\"{i}\"><c:v>{caption}{i + 1}</c:v></c:pt>"));
        string values = string.Concat(Enumerable.Range(0, count).Select(i => $"<c:pt idx=\"{i}\"><c:v>{2 + i % 3}</c:v></c:pt>"));
        string plot = area ? "area" : "line";
        string series = string.Concat((legendNames ?? ["Revenue"]).Select((name, index) =>
            $"<c:ser><c:idx val=\"{index}\"/><c:order val=\"{index}\"/><c:tx><c:v>{name}</c:v></c:tx>" +
            $"<c:cat><c:strLit>{categories}</c:strLit></c:cat><c:val><c:numLit>{values}</c:numLit></c:val><c:smooth val=\"0\"/></c:ser>"));
        string legendBody = legendFontSize is { } size
            ? $"<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"{size * 100}\"><a:latin typeface=\"LegendFace\"/></a:defRPr></a:pPr><a:endParaRPr/></a:p></c:txPr>" : "";
        string legendShape = legendFrame ? "<c:spPr><a:solidFill><a:srgbClr val=\"FFF5E0\"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>" : "";
        string manual = fallback == "manual" ? "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\"0.15\"/><c:y val=\"0.1\"/><c:w val=\"0.65\"/><c:h val=\"0.7\"/></c:manualLayout></c:layout>" : "";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea>{{manual}}<c:{{plot}}Chart><c:varyColors val="0"/>
                {{series}}<c:axId val="1"/><c:axId val="2"/></c:{{plot}}Chart>
                <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/></c:scaling>
                  <c:axPos val="b"/><c:crossAx val="{{(fallback == "unpaired" ? "other" : "2")}}"/><c:crosses val="autoZero"/>
                  <c:txPr><a:bodyPr rot="{{(fallback == "rotated" ? "-2700000" : "0")}}" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{{fontSize * 100}}"/></a:pPr><a:endParaRPr/></a:p></c:txPr>
                </c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling>{{(fallback == "log" ? "<c:logBase val=\"10\"/>" : "")}}<c:min val="1"/><c:max val="5"/></c:scaling><c:axPos val="l"/>
                  <c:tickLblPos val="{{side}}"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="{{(between ? "between" : "midCat")}}"/><c:majorUnit val="1"/>
                </c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/>{{legendBody}}{{legendShape}}</c:legend></c:chart>
            </c:chartSpace>
            """;
        const string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="457200"/><a:ext cx="9144000" cy="5486400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(resolver ?? new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }

    private sealed class LegendLineBoxResolver : IFontResolver
    {
        private readonly IFontResolver other = new PortableRenderingTests.PinnedResolver();
        private readonly IFontProgramSource source;

        internal LegendLineBoxResolver(ushort ascender)
        {
            byte[] bytes = TestFontBuilder.CreateTestFont();
            var os2 = TestFontBuilder.GetTableRange(bytes, "OS/2");
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(os2.Offset + 74, 2), ascender);
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(os2.Offset + 76, 2), 200);
            source = new MemoryFontProgramSource("test:legend-line-box:" + ascender, bytes);
        }

        public FontFaceResolution Resolve(FontRequest request) => request.FamilyName == "LegendFace"
            ? new(request.FamilyName, request.FamilyName, new FontStyleKey(request.Bold, request.Italic), source, IsFallback: false)
            : other.Resolve(request);
    }
}
