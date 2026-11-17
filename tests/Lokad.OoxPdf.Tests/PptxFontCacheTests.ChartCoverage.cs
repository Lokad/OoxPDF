using System.Buffers.Binary;
using System.Reflection;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxFontCacheTests
{
    public static void PptxChartCoverageFallbackKeepsShapeTitleAndLegendContent()
    {
        foreach (string kind in new[] { "partial", "full", "cff", "unresolved", "all-fallback" })
        {
            byte[] primaryBytes = kind switch
            {
                "full" => CoverageFont(),
                "cff" => TestFontBuilder.CreateCffKindFont("Primary"),
                _ => TestFontBuilder.CreateTestFont()
            };
            IFontResolver resolver = kind == "unresolved"
                ? new SingleProgramResolver(new MemoryFontProgramSource("coverage-primary", primaryBytes))
                : new CatalogResolver(FontFace("Primary", "coverage-primary", primaryBytes),
                    FontFace("Fallback", "coverage-fallback", CoverageFont()));
            var parts = CoverageParts();
            if (kind == "all-fallback") parts["ppt/charts/chart1.xml"] = parts["ppt/charts/chart1.xml"].Replace("Title AŁB", "Ł", StringComparison.Ordinal);
            using var source = TestFixtures.CreateZipPackage(parts);
            using var package = OoxPackage.Open(source, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = new PptxRenderer(resolver).RenderPagesAsync(new PptxReader().Read(package, default),
                package, diagnostics.Add, default).ToBlockingEnumerable().Single();
            string value = kind == "unresolved" ? "A?B" : "A\u0141B";
            foreach (var owner in new[] { (Prefix: "F", Label: "Shape "), (Prefix: "CT", Label: "Title "), (Prefix: "CL", Label: "Series ") })
            {
                string painted = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                    fontResourceFilter: name => name.StartsWith(owner.Prefix, StringComparison.Ordinal)).Select(show => show.Text));
                TestAssert.Equal(kind == "all-fallback" && owner.Prefix == "CT" ? "Ł" : owner.Label + value, painted);
            }
            TestAssert.Equal("https://example.invalid/coverage", page.Annotations.Single().Uri);
            TestAssert.True(page.Annotations[0].Width > 0 && page.Annotations[0].Height > 0, "The linked title retains its measured frame.");
            TestAssert.True((kind == "unresolved") == diagnostics.Any(d => d.Id == "FONT_MISSING_GLYPHS"),
                $"{kind} coverage: only genuinely unavailable glyphs receive a substitution diagnostic.");
            if (kind == "unresolved") TestAssert.Equal("Primary", diagnostics.Single(d => d.Id == "FONT_MISSING_GLYPHS").Feature);
            TestAssert.Equal(kind == "cff", diagnostics.Any(d => d.Id == "FONT_UNSUPPORTED_OUTLINES"));
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_NODE_RENDER_FAILED"), "Font fallback must preserve the complete chart.");
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    public static void PptxChartCoverageSplitsKeepMeasuredAlignmentTrackingAndStyle()
    {
        var split = typeof(PptxRenderer).GetMethod("SplitRunsByResolvedTypefaceAsync", BindingFlags.NonPublic | BindingFlags.Static)!;
        OpenTypeFont full = OpenTypeFont.Load(CoverageFont());
        foreach (PptxRenderer.TextAlignment alignment in new[] { PptxRenderer.TextAlignment.Left, PptxRenderer.TextAlignment.Center, PptxRenderer.TextAlignment.Right })
        foreach (double tracking in new[] { 0d, 1.25d })
        {
            var run = new PptxRenderer.TextRun("A\u0141B", 20, 100, 90, 24, 0, 0, 200, 200, 12, tracking, 2,
                new RgbColor(24, 80, 160), .6, new RgbColor(255, 255, 0), false, false, true, true, false,
                alignment, "Primary", 0, 0, 0, false, false, false, null, false);
            var resolver = new PresentationFontResolver(new CatalogResolver(
                FontFace("Primary", "coverage-primary", TestFontBuilder.CreateTestFont()),
                FontFace("Fallback", "coverage-fallback", CoverageFont())));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var result = ((ValueTask<List<PptxRenderer.TextRun>>)split.Invoke(null,
                [new[] { run }, resolver, (Action<OoxPdfDiagnostic>)diagnostics.Add, CancellationToken.None])!).GetAwaiter().GetResult();
            TestAssert.Equal(3, result.Count);
            TestAssert.True(result.Select(r => r.FontFamily).SequenceEqual(new[] { "Primary", "Fallback", "Primary" }), "Every glyph uses the face selected for measurement.");
            double[] advances = "A\u0141B".Select(c => full.GetAdvanceWidth(full.MapCodePoint(c)) * 12d / full.UnitsPerEm).ToArray();
            double natural = advances.Sum() + 2 * tracking;
            double x = run.X + (alignment == PptxRenderer.TextAlignment.Center ? (run.Width - natural) / 2 : alignment == PptxRenderer.TextAlignment.Right ? run.Width - natural : 0);
            for (int index = 0; index < result.Count; index++)
            {
                var part = result[index];
                TestAssert.True(Math.Abs(part.X - x) < 1e-9 && Math.Abs(part.Width - advances[index]) < 1e-9, "Alignment and inter-face tracking retain the measured glyph origins.");
                TestAssert.Equal(run with { Text = run.Text[index].ToString(), X = part.X, Width = part.Width, FontFamily = part.FontFamily,
                    Alignment = PptxRenderer.TextAlignment.Left, PreventCoalesce = true }, part);
                x += advances[index] + tracking;
            }
            TestAssert.Equal(0, diagnostics.Count);
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Equal(cancelled.Token, TestAssert.Throws<OperationCanceledException>(() =>
                ((ValueTask<List<PptxRenderer.TextRun>>)split.Invoke(null,
                    [new[] { run }, resolver, null, cancelled.Token])!).GetAwaiter().GetResult()).CancellationToken);
        }
    }

    // The primary retains ASCII. Add one independently mapped fallback character
    // to the test font's format-12 group; real outline/ink evidence uses portable fonts.
    private static byte[] CoverageFont()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        int cmap = TestFontBuilder.GetTableRange(bytes, "cmap").Offset;
        int table = cmap + (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(cmap + 16));
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(table + 28), 0x141);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(table + 32), 0x141);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(table + 36), 66);
        return bytes;
    }

    private static Dictionary<string, string> CoverageParts() => new()
    {
        ["[Content_Types].xml"] = PptxTests.BasicContentTypes().Replace("</Types>", "<Override PartName=\"/ppt/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/></Types>"),
        ["_rels/.rels"] = PptxTests.PackageRelationship(),
        ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
        ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
        ["ppt/slides/slide1.xml"] = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape control"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="254000"/><a:ext cx="5486400" cy="635000"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Primary"/></a:rPr><a:t>Shape AŁB</a:t></a:r></a:p></p:txBody></p:sp>
            <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1270000"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="chart"/></a:graphicData></a:graphic></p:graphicFrame>
            </p:spTree></p:cSld></p:sld>
            """,
        ["ppt/slides/_rels/slide1.xml.rels"] = """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>
            """,
        ["ppt/charts/chart1.xml"] = """
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>
            <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800" spc="125" u="sng" strike="sngStrike"><a:latin typeface="Primary"/><a:hlinkClick r:id="link"/></a:rPr><a:t>Title AŁB</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:plotArea><c:scatterChart><c:scatterStyle val="marker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strLit><c:pt idx="0"><c:v>Series AŁB</c:v></c:pt></c:strLit></c:tx><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser></c:scatterChart></c:plotArea>
            <c:legend><c:legendPos val="b"/></c:legend></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" spc="125"><a:latin typeface="Primary"/></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>
            """,
        ["ppt/charts/_rels/chart1.xml.rels"] = """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="link" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/coverage" TargetMode="External"/></Relationships>
            """
    };
}
