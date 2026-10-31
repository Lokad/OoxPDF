using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxLegendDefaultTests
{
    public static void PptxOmittedLegendOverlayMatchesPowerPointDefaultAcrossKinds()
    {
        foreach (string kind in new[] { "bar", "column", "line", "pie", "doughnut" })
        {
            PdfPage explicitOverlay = Render(kind, "1");
            foreach (string? flag in new string?[] { null, "" })
            {
                PdfPage implicitOverlay = Render(kind, flag);
                TestAssert.Equal(explicitOverlay.Content, implicitOverlay.Content);
                TestAssert.True(DocxCommentsTests.ReadEmbeddedGlyphTextShows(implicitOverlay)
                    .Any(show => show.Text.Contains("Series", StringComparison.Ordinal) || show.Text.Contains("Q1", StringComparison.Ordinal)),
                    "An implicit overlay preserves the supported legend text.");
            }
        }
    }

    public static void PptxExplicitReservedLegendKeepsItsDistinctPlotContract()
    {
        foreach (string kind in new[] { "bar", "column" })
        {
            PdfPage reserved = Render(kind, "0");
            PdfPage overlaid = Render(kind, "1");
            TestAssert.True(reserved.Content != overlaid.Content,
                "An explicitly reserved legend keeps its separate plot-space contract.");
        }
    }

    private static PdfPage Render(string kind, string? overlay)
    {
        const string series = "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:pt idx=\"0\"><c:v>Q1</c:v></c:pt><c:pt idx=\"1\"><c:v>Q2</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>3</c:v></c:pt><c:pt idx=\"1\"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>";
        bool axes = kind is "bar" or "column" or "line";
        string node = kind is "bar" or "column" ? "barChart" : kind + "Chart";
        string direction = kind is "bar" or "column" ? "<c:barDir val=\"" + (kind == "bar" ? "bar" : "col") + "\"/>" : "";
        string plot = "<c:" + node + ">" + direction + series + (axes ? "<c:axId val=\"1\"/><c:axId val=\"2\"/>" : "") + "</c:" + node + ">";
        string axisXml = axes ? "<c:catAx><c:axId val=\"1\"/><c:axPos val=\"" + (kind == "bar" ? "l" : "b") + "\"/><c:crossAx val=\"2\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling><c:min val=\"0\"/><c:max val=\"5\"/></c:scaling><c:axPos val=\"" + (kind == "bar" ? "b" : "l") + "\"/><c:crossAx val=\"1\"/></c:valAx>" : "";
        string flag = overlay is null ? "" : overlay.Length == 0 ? "<c:overlay/>" : "<c:overlay val=\"" + overlay + "\"/>";
        string chart = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea>" + plot + axisXml + "</c:plotArea><c:legend><c:legendPos val=\"b\"/>" + flag + "</c:legend></c:chart></c:chartSpace>";
        using var stream = PptxTests.CreateSingleChartPackageStream(chart);
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }
}
