using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxTableMergeTests
{
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void PptxMergedTableNeighborsKeepAllSourceContent()
    {
        using JsonDocument golden = JsonDocument.Parse(File.ReadAllText(Case("office.json")));
        foreach (JsonElement control in golden.RootElement.GetProperty("Cases").EnumerateArray())
        {
            string input = Case(control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var (document, package) = Read(input);
            PdfPage page = new PptxRenderer(new TestFaceFontResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            string painted = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
            foreach (JsonElement cell in control.GetProperty("Cells").EnumerateArray())
                TestAssert.Contains(cell.GetProperty("Text").GetString()!, painted);
        }
    }

    public static void PptxMergedTableColumnsKeepStyleAndLinkOwnership()
    {
        var (document, package) = Read(Case("styled-body.pptx"));
        PptxSceneTable table = new PptxSceneBuilder().Build(document, package, default).Slides[0].SlideNodes
            .Select(node => node.Table).OfType<PptxSceneTable>().Single();
        foreach (PptxSceneTableRow row in table.Rows)
            for (int column = 0; column < row.Cells.Count; column++)
                TestAssert.Equal(column, row.Cells[column].ColumnIndex);
        TestAssert.True(table.Rows[1].Cells[2].StyleText.Bold, "The following cell must retain the last-column style.");
        TestAssert.Equal(new RgbColor(255, 255, 255), table.Rows[1].Cells[2].StyleText.Color!.Value);
        PptxTextFrameModelSnapshot frame = TailFrame(document, package);
        TestAssert.Equal(2, frame.TableColumnIndex!.Value);
        Near(360, frame.FrameX, .000001, "The cell starts at its source grid column.");
        Near(144, frame.FrameWidth, .000001, "The cell keeps its source grid width.");

        foreach (bool compact in new[] { false, true })
            WithSlide("merged", xml => {
                XElement row = xml.Descendants(A + "tr").First();
                if (compact) row.Elements(A + "tc").ElementAt(1).Remove();
                TailCell(xml).Descendants(A + "rPr").First().Add(new XElement(A + "hlinkClick", new XAttribute(R + "id", "rIdMergeLink")));
            }, path => {
                var source = Read(path);
                PdfPage page = new PptxRenderer(new TestFaceFontResolver()).RenderPagesAsync(source.Document, source.Package, null, default).ToBlockingEnumerable().Single();
                PdfLinkAnnotation link = page.Annotations.Single();
                TestAssert.Equal("https://example.invalid/merged-cell", link.Uri);
                TestAssert.True(link.Width > 0 && link.X >= 360 && link.X + link.Width <= 504,
                    "Dense and compact merge forms must retain the following cell's visible link bounds.");
                TestAssert.Contains("Right tail", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text)));
            }, parts => {
                const string name = "ppt/slides/_rels/slide1.xml.rels";
                XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
                XDocument xml = parts.TryGetValue(name, out byte[]? bytes) ? XDocument.Parse(Encoding.UTF8.GetString(bytes)) : new(new XElement(rel + "Relationships"));
                xml.Root!.Add(new XElement(rel + "Relationship", new XAttribute("Id", "rIdMergeLink"),
                    new XAttribute("Type", R.NamespaceName + "/hyperlink"), new XAttribute("Target", "https://example.invalid/merged-cell"), new XAttribute("TargetMode", "External")));
                parts[name] = Encoding.UTF8.GetBytes(xml.ToString());
            });
    }

    public static void PptxMergedTableHeightMatchesCompactRecovery()
    {
        double? firstHeight = null;
        foreach (bool compact in new[] { false, true })
            WithSlide("merged", xml => {
                if (compact) xml.Descendants(A + "tr").First().Elements(A + "tc").ElementAt(1).Remove();
                // Exercise the existing content-minimum path; a tightly declared
                // frame deliberately retains its fixed rows instead of expanding.
                xml.Descendants(P + "xfrm").Single().Element(A + "ext")!.SetAttributeValue("cy", 3048000);
                XElement paragraph = TailCell(xml).Descendants(A + "p").First();
                XElement runProperties = new(paragraph.Descendants(A + "rPr").First());
                paragraph.RemoveNodes();
                for (int index = 1; index <= 10; index++)
                {
                    if (index != 1) paragraph.Add(new XElement(A + "br"));
                    paragraph.Add(new XElement(A + "r", new XElement(runProperties), new XElement(A + "t", $"Right tail {index}")));
                }
            }, path => {
                var source = Read(path);
                double height = TailFrame(source.Document, source.Package).FrameHeight;
                TestAssert.True(height > 72, "The following cell's ten lines must participate in row-height measurement.");
                if (firstHeight is double expected) Near(expected, height, .000001, "Continuation placeholders must not change measured row height.");
                firstHeight = height;
            });
    }

    public static void PptxMergedTableGeometryMatchesIndependentOfficeOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent Office geometry requires the source-declared installed Aptos faces.");
        var fonts = new WindowsFontResolver();
        foreach (bool bold in new[] { false, true })
        {
            FontFaceResolution face = fonts.Resolve(new FontRequest("Aptos", bold, false));
            if (face.IsFallback || !face.ResolvedFamily.Equals("Aptos", StringComparison.OrdinalIgnoreCase))
                TestAssert.Skip("Independent Office geometry requires the source-declared installed Aptos faces.");
        }
        using JsonDocument golden = JsonDocument.Parse(File.ReadAllText(Case("office.json")));
        foreach (JsonElement control in golden.RootElement.GetProperty("Cases").EnumerateArray())
        {
            string input = Case(control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var source = Read(input);
            PdfPage page = new PptxRenderer(fonts).RenderPagesAsync(source.Document, source.Package, null, default).ToBlockingEnumerable().Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            foreach (JsonElement cell in control.GetProperty("Cells").EnumerateArray())
            {
                string text = cell.GetProperty("Text").GetString()!;
                var matches = shows.Where(show => show.Text == text).ToArray();
                TestAssert.True(matches.Length == 1, "Expected one painted source cell: " + text);
                var actual = matches.Single();
                Near(cell.GetProperty("OfficeX").GetDouble(), actual.X, golden.RootElement.GetProperty("OriginTolerancePoints").GetDouble(), "Independent Office cell X: " + text);
                Near(cell.GetProperty("OfficeTopBaseline").GetDouble(), page.Height - actual.Y, golden.RootElement.GetProperty("OriginTolerancePoints").GetDouble(), "Independent Office cell Y: " + text);
                Near(cell.GetProperty("OfficeFontSize").GetDouble(), actual.Size, golden.RootElement.GetProperty("FontSizeTolerancePoints").GetDouble(), "Independent Office cell size: " + text);
            }
        }
    }

    private static PptxTextFrameModelSnapshot TailFrame(PptxDocument document, OoxPackage package) =>
        PptxRenderer.InspectTableTextFrameModels(document, package, 0, inspectionFonts: new PresentationFontResolver(new TestFaceFontResolver()))
            .Single(frame => frame.Paragraphs.SelectMany(paragraph => paragraph.Runs).Any(run => run.Text.StartsWith("Right tail", StringComparison.Ordinal)));

    private static XElement TailCell(XDocument xml) => xml.Descendants(A + "tc")
        .Single(cell => cell.Descendants(A + "t").Any(text => text.Value == "Right tail"));

    private static (PptxDocument Document, OoxPackage Package) Read(string path)
    {
        OoxPackage package = OoxPackage.Open(path, default);
        return (new PptxReader().Read(package, default), package);
    }

    private static void WithSlide(string id, Action<XDocument> change, Action<string> check, Action<Dictionary<string, byte[]>>? changeParts = null)
    {
        using var archive = ZipFile.OpenRead(Case(id + ".pptx"));
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var source = entry.Open(); using var output = new MemoryStream(); source.CopyTo(output); return output.ToArray();
        });
        const string slide = "ppt/slides/slide1.xml";
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(parts[slide]));
        change(xml); parts[slide] = Encoding.UTF8.GetBytes(xml.ToString());
        changeParts?.Invoke(parts);
        string path = TestFixtures.WriteTempPackage(".pptx", parts);
        try { check(path); } finally { File.Delete(path); }
    }

    private static void Near(double expected, double actual, double tolerance, string message) =>
        TestAssert.True(double.IsFinite(actual) && Math.Abs(actual - expected) <= tolerance,
            $"{message} Expected {expected}, actual {actual}, tolerance {tolerance}.");

    private static string Case(string file) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "TableMerges", file));
}
