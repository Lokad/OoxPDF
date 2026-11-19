using System.Globalization;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAffineGeometryTests
{
    public static void PptxNodeHyperlinkBoundsFollowLeafPaintTransforms()
    {
        foreach (var sample in new[] {
            (Angle: 0, Flip: false, Expected: new PdfRectangle(90, 425, 120, 30)),
            (Angle: 90, Flip: false, Expected: new PdfRectangle(135, 380, 30, 120)),
            (Angle: 45, Flip: true, Expected: new PdfRectangle(96.966991411, 386.966991411, 106.066017178, 106.066017178))
        })
        {
            var annotation = RenderAffineLink(sample.Angle, sample.Flip, grouped: false);
            AssertFrame(sample.Expected, annotation);
        }
    }

    public static void PptxNodeHyperlinkBoundsComposeGroupScaleAndLeafRotation()
    {
        // The group doubles child X/width and translates by (36,54) points.
        // The 240x30 leaf then rotates about its own center (336,386 in PDF).
        var annotation = RenderAffineLink(90, false, grouped: true);
        AssertFrame(new PdfRectangle(321, 266, 30, 240), annotation);
    }

    public static void PptxReflectedGroupPaintMatchesFlattenedTriangle()
    {
        // The 90-degree outer turn takes the reflected leaf center around (216,174).
        // An odd number of reflected axes reverses the 90-degree leaf turn.
        foreach (var sample in new[] {
            (FlipH: 1, FlipV: 0, X: 176, Y: 189, Angle: 0),
            (FlipH: 0, FlipV: 1, X: 136, Y: 129, Angle: 0),
            (FlipH: 0, FlipV: 0, X: 176, Y: 129, Angle: 180),
            (FlipH: 1, FlipV: 1, X: 136, Y: 189, Angle: 180)
        })
        {
            string child = Triangle(90, 85, 90);
            string grouped = Group(3, 36, 54, 360, 240, 90, sample.FlipH, sample.FlipV, child);
            string flat = Triangle(sample.X, sample.Y, sample.Angle, sample.FlipH, sample.FlipV);
            AssertSameTriangle(flat, grouped);
        }

        // Inner vertical reflection and turn place the leaf center at (155,95)
        // in the outer group. The outer horizontal reflection places it at (241,149)
        // on the slide and reverses the inner turn.
        string inner = Group(4, 60, 45, 240, 160, 90, 0, 1, Triangle(30, 40, 0));
        AssertSameTriangle(Triangle(181, 134, 270, 1, 1), Group(3, 36, 54, 360, 240, 0, 1, 0, inner));

        static void AssertSameTriangle(string flat, string grouped)
        {
            PdfPage expected = RenderLinkPage(flat);
            PdfPage actual = RenderLinkPage(grouped);
            TestAssert.Contains("0.855 0.463 0.165 rg", expected.Content);
            TestAssert.Equal(expected.Content, actual.Content);
            TestAssert.Equal(expected.Annotations.Single(), actual.Annotations.Single());
        }

        static string Triangle(int x, int y, int angle, int flipH = 0, int flipV = 0) => $$"""
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Linked triangle"><a:hlinkClick r:id="link"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm rot="{{angle * 60000}}" flipH="{{flipH}}" flipV="{{flipV}}"><a:off x="{{x * 12700}}" y="{{y * 12700}}"/><a:ext cx="1524000" cy="381000"/></a:xfrm>
              <a:prstGeom prst="rtTriangle"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DA762A"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>
            """;

        static string Group(int id, int x, int y, int width, int height, int angle, int flipH, int flipV, string child) => $$"""
            <p:grpSp><p:nvGrpSpPr><p:cNvPr id="{{id}}" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm rot="{{angle * 60000}}" flipH="{{flipH}}" flipV="{{flipV}}"><a:off x="{{x * 12700}}" y="{{y * 12700}}"/><a:ext cx="{{width * 12700}}" cy="{{height * 12700}}"/>
              <a:chOff x="0" y="0"/><a:chExt cx="{{width * 12700}}" cy="{{height * 12700}}"/></a:xfrm></p:grpSpPr>{{child}}</p:grpSp>
            """;
    }

    public static void PptxUnrotatedFramePaintersKeepLinkFallback()
    {
        foreach (bool chart in new[] { false, true })
        {
            string data = chart
                ? "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"chart\"/></a:graphicData>"
                : "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblGrid><a:gridCol w=\"1524000\"/></a:tblGrid><a:tr h=\"381000\"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData>";
            Dictionary<string, string>? extra = chart ? new() {
                ["ppt/charts/chart1.xml"] = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea><c:pieChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:val><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"
            } : null;
            PdfPage Render(int angle) => RenderLinkPage($"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"2\" name=\"Frame\"><a:hlinkClick r:id=\"link\"/></p:cNvPr><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm rot=\"{angle * 60000}\"><a:off x=\"1143000\" y=\"1079500\"/><a:ext cx=\"1524000\" cy=\"381000\"/></p:xfrm><a:graphic>{data}</a:graphic></p:graphicFrame>", extra,
                chart ? "<Relationship Id=\"chart\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart1.xml\"/>" : "");
            var ordinary = Render(0);
            var rotated = Render(90);
            TestAssert.Equal(ordinary.Content, rotated.Content);
            AssertFrame(new PdfRectangle(90, 425, 120, 30), rotated.Annotations.Single());
        }
    }

    public static void PptxRotatedTextPaintKeepsDeclaredDeviceClip()
    {
        var prepare = typeof(PptxRenderer).GetMethod("PrepareTextPaint",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected shared text paint setup.");
        foreach (var sample in new[] {
            (Angle: 0d, FlipH: false, FlipV: false, Left: 10d, Bottom: 20d),
            (Angle: -90d, FlipH: false, FlipV: false, Left: 10d, Bottom: 20d),
            (Angle: 45d, FlipH: false, FlipV: false, Left: 10d, Bottom: 20d),
            (Angle: 45d, FlipH: true, FlipV: false, Left: 10d, Bottom: 20d),
            // Unrotated flips retain their existing local-clip convention.
            (Angle: 0d, FlipH: true, FlipV: false, Left: 160d, Bottom: 20d),
            (Angle: 0d, FlipH: false, FlipV: true, Left: 10d, Bottom: 340d)
        })
        {
            var run = new PptxRenderer.TextRun("", 0, 0, 0, 0, 10, 20, 30, 40, 12, 0, 0,
                new RgbColor(0, 0, 0), 1, null, false, false, false, false, true,
                PptxRenderer.TextAlignment.Center, null, sample.Angle, 100, 200, sample.FlipH, sample.FlipV, false, null, false);
            var graphics = new PdfGraphicsBuilder();
            graphics.SaveState(); prepare.Invoke(null, [graphics, run]); graphics.RestoreState();
            string[] lines = graphics.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            int clipAt = Array.FindIndex(lines, line => line.EndsWith(" re W* n", StringComparison.Ordinal));
            int matrixAt = Array.FindIndex(lines, line => line.EndsWith(" cm", StringComparison.Ordinal));
            double[] clip = lines[clipAt].Split(' ').Take(4).Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            double[] matrix = matrixAt >= 0 && matrixAt < clipAt
                ? lines[matrixAt].Split(' ').Take(6).Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray()
                : [1, 0, 0, 1, 0, 0];
            var expected = new[] { (sample.Left, sample.Bottom), (sample.Left + 30, sample.Bottom),
                (sample.Left + 30, sample.Bottom + 40), (sample.Left, sample.Bottom + 40) };
            foreach (var point in new[] { (clip[0], clip[1]), (clip[0] + clip[2], clip[1]),
                (clip[0] + clip[2], clip[1] + clip[3]), (clip[0], clip[1] + clip[3]) })
            {
                double x = matrix[0] * point.Item1 + matrix[2] * point.Item2 + matrix[4];
                double y = matrix[1] * point.Item1 + matrix[3] * point.Item2 + matrix[5];
                TestAssert.True(expected.Any(corner => Math.Abs(corner.Item1 - x) < .001 && Math.Abs(corner.Item2 - y) < .001),
                    "Every emitted clip corner must match the declared device frame, including oblique rotation.");
            }
        }
    }

    private static void AssertFrame(PdfRectangle expected, PdfLinkAnnotation actual)
    {
        TestAssert.Equal("https://example.invalid/node", actual.Uri);
        TestAssert.True(Math.Abs(expected.X - actual.X) < .001 && Math.Abs(expected.Y - actual.Y) < .001 &&
            Math.Abs(expected.Width - actual.Width) < .001 && Math.Abs(expected.Height - actual.Height) < .001,
            $"Node link must cover the independently specified transformed frame {expected}; got {actual}.");
    }

    private static PdfLinkAnnotation RenderAffineLink(int angle, bool flip, bool grouped)
    {
        string shape = $$"""
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Linked frame"><a:hlinkClick r:id="link"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm rot="{{angle * 60000}}" flipH="{{(flip ? "1" : "0")}}"><a:off x="1143000" y="1079500"/><a:ext cx="1524000" cy="381000"/></a:xfrm>
              <a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>
            """;
        if (grouped) shape = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"3\" name=\"Group\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>" +
            "<a:off x=\"457200\" y=\"685800\"/><a:ext cx=\"6096000\" cy=\"3810000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"3048000\" cy=\"3810000\"/>" +
            "</a:xfrm></p:grpSpPr>" + shape + "</p:grpSp>";
        return RenderLinkPage(shape).Annotations.Single();
    }

    private static PdfPage RenderLinkPage(string shape, Dictionary<string, string>? extra = null, string extraRelationships = "")
    {
        var parts = new Dictionary<string, string> {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/slides/_rels/slide1.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"link\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.invalid/node\" TargetMode=\"External\"/>" + extraRelationships + "</Relationships>",
            ["ppt/slides/slide1.xml"] = $"<p:sld xmlns:p=\"{OoxNamespaces.PresentationNamespace}\" xmlns:a=\"{OoxNamespaces.DrawingNamespace}\" xmlns:r=\"{OoxNamespaces.RelationshipsNamespace}\"><p:cSld><p:spTree>{shape}</p:spTree></p:cSld></p:sld>"
        };
        if (extra is not null) foreach (var part in extra) parts.Add(part.Key, part.Value);
        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var document = new PptxReader().Read(package, default, diagnostics.Add);
        var page = new PptxRenderer(new TestFaceFontResolver()).RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().Single();
        TestAssert.Equal(0, diagnostics.Count);
        return page;
    }
}
