using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLegendsTests
{
    public static void PptxAutomaticHorizontalLegendsPaintCompleteSourceNames()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "ComposedAxes", "shared-column-line.pptx"));
        using var archive = ZipFile.OpenRead(path);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var source = entry.Open(); using var bytes = new MemoryStream();
            source.CopyTo(bytes); return bytes.ToArray();
        });
        XNamespace c = OoxNamespaces.ChartNamespace, a = OoxNamespaces.DrawingNamespace;
        string[] names = ["Monthly planning", "Yearly progress"];
        foreach (var (size, side) in new[] { (12, "b"), (18, "b"), (24, "t") })
        foreach (bool composed in new[] { true, false })
        {
            var chart = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["ppt/charts/chart1.xml"]));
            var plot = chart.Descendants(c + "plotArea").Single();
            var column = plot.Element(c + "barChart")!;
            var line = plot.Element(c + "lineChart")!;
            var series = new[] { column.Element(c + "ser")!, line.Element(c + "ser")! };
            for (int index = 0; index < series.Length; index++)
                series[index].Element(c + "tx")!.ReplaceWith(new XElement(c + "tx", new XElement(c + "v", names[index])));
            if (!composed)
            {
                // Keep two names and their data while removing the plot-kind boundary.
                var second = new XElement(series[0]);
                second.Element(c + "idx")!.SetAttributeValue("val", 1);
                second.Element(c + "order")!.SetAttributeValue("val", 1);
                second.Element(c + "tx")!.ReplaceWith(new XElement(series[1].Element(c + "tx")!));
                second.Element(c + "val")!.ReplaceWith(new XElement(series[1].Element(c + "val")!));
                series[0].AddAfterSelf(second);
                line.Remove();
            }
            plot.AddAfterSelf(new XElement(c + "legend", new XElement(c + "legendPos", new XAttribute("val", side)),
                new XElement(c + "layout"), new XElement(c + "overlay", new XAttribute("val", 0)),
                new XElement(c + "txPr", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"),
                    new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", size * 100),
                        new XElement(a + "latin", new XAttribute("typeface", "DejaVu Sans")))), new XElement(a + "endParaRPr")))));
            var input = new Dictionary<string, byte[]>(parts)
            {
                ["ppt/charts/chart1.xml"] = TestFixtures.Utf8(chart.ToString(SaveOptions.DisableFormatting))
            };
            using var stream = TestFixtures.CreateZipPackage(input);
            using var package = OoxPackage.Open(stream, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default)
                .ToBlockingEnumerable().Single();
            foreach (string name in names)
                PortableRenderingTests.ReadPaintedSourceLabel(page, name, fontPrefix: "CL");
        }
    }
}
