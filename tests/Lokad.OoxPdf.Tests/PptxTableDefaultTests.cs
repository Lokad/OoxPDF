using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxTableDefaultTests
{
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly (string Id, bool Bold, RgbColor Color)[] Controls =
    [
        ("inherited", true, new(255,255,255)),
        ("run-clear-red-valid", false, new(255,0,0)),
        ("paragraph-clear-red", false, new(255,0,0)),
        ("paragraph-clear", false, new(255,255,255)),
        ("paragraph-red", true, new(255,0,0))
    ];

    public static void PptxTableParagraphDefaultsMatchIndependentOfficeControls()
    {
        foreach (var control in Controls)
            CheckPaint(Case(control.Id), control.Bold, control.Color);
    }

    public static void PptxTableParagraphDefaultsRetainDirectRunAndAbsentPropertyFallbacks()
    {
        WithSlide("paragraph-clear-red", xml => {
            XElement properties = xml.Descendants(A+"tc").First().Descendants(A+"rPr").First();
            properties.SetAttributeValue("b", "1");
            properties.AddFirst(new XElement(A+"solidFill", new XElement(A+"srgbClr", new XAttribute("val", "0000FF"))));
        }, path => { CheckPaint(path, true, new(0,0,255)); });
        WithSlide("inherited", xml => {
            XElement paragraph = xml.Descendants(A+"tc").First().Descendants(A+"p").First();
            paragraph.AddFirst(new XElement(A+"pPr", new XElement(A+"defRPr")));
        }, path => { CheckPaint(path, true, new(255,255,255)); });
    }

    public static void PptxTableParagraphDefaultsKeepSourceXmlAndCancellation()
    {
        using var input = File.OpenRead(Case("paragraph-clear-red"));
        OoxPackage package = OoxPackage.Open(input, default);
        PptxDocument document = new PptxReader().Read(package, default);
        XDocument xml = package.LoadXml(package.GetPart(document.Slides[0].PartName)!, default);
        string before = xml.ToString(SaveOptions.DisableFormatting);
        var renderer = new PptxRenderer(new PortableRenderingTests.PinnedResolver());
        string first = renderer.RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single().Content;
        string second = renderer.RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single().Content;
        TestAssert.Equal(first, second);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, package, null, cancelled.Token).ToBlockingEnumerable().ToArray());
        TestAssert.Equal(before, xml.ToString(SaveOptions.DisableFormatting));
    }

    private static void CheckPaint(string path, bool bold, RgbColor color)
    {
        using var input = File.OpenRead(path);
        OoxPackage package = OoxPackage.Open(input, default);
        PptxDocument document = new PptxReader().Read(package, default);
        PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
        // The first text object paints the header; the body has its own bold
        // inherited formatting. The pinned face paints requested bold with stroke.
        int start = page.Content.IndexOf("BT" + Environment.NewLine, StringComparison.Ordinal);
        TestAssert.True(start >= 0, "Expected the header text object.");
        int end = page.Content.IndexOf(Environment.NewLine + "ET", start, StringComparison.Ordinal);
        TestAssert.True(end > start, "Expected a complete header text object.");
        string header = page.Content[start..end];
        TestAssert.Equal(bold, header.Contains("2 Tr", StringComparison.Ordinal));
        string fill = color == new RgbColor(255,255,255) ? "1 g"
            : color == new RgbColor(0,0,255) ? "0 0 1 rg" : "1 0 0 rg";
        TestAssert.Contains("BT" + Environment.NewLine + fill + Environment.NewLine, header);
        TestAssert.True(page.Fonts.Count > 0, "Supported cell text retains its font resources.");
    }

    private static string Case(string id) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "TableDefaults", "table-"+id+".pptx"));

    private static void WithSlide(string id, Action<XDocument> change, Action<string> check)
    {
        using var archive = ZipFile.OpenRead(Case(id));
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var input = entry.Open(); using var bytes = new MemoryStream(); input.CopyTo(bytes); return bytes.ToArray();
        });
        XDocument slide = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        change(slide); parts["ppt/slides/slide1.xml"] = System.Text.Encoding.UTF8.GetBytes(slide.ToString());
        string path = TestFixtures.WriteTempPackage(".pptx", parts);
        try { check(path); }
        finally { File.Delete(path); }
    }
}
