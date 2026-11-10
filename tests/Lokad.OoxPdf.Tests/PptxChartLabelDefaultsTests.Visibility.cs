using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    // Independent Office controls cover these five families.
    // A key is an adornment of label text, not a standalone label unit.
    private static readonly (string Input, int Keys, bool Stroke)[] VisibilityCases =
    [
        ("BuiltinPalettes/line-3.pptx", 18, true),
        ("BuiltinPalettes/area-3.pptx", 18, false),
        ("BuiltinPalettes/count-series-3.pptx", 18, false),
        ("BuiltinPalettes/pie-vary.pptx", 6, false),
        ("pptx-ladder-11-chart-scatter-clusters-port.pptx", 10, false)
    ];

    public static void PptxEmptyDataLabelKeysPreserveDisabledContentAcrossPlotKinds()
    {
        foreach (var source in VisibilityCases)
        {
            var disabled = RenderVisibilityCase(source.Input, key: false, value: false);
            var keyOnly = RenderVisibilityCase(source.Input, key: true, value: false);
            TestAssert.Equal(disabled.Content, keyOnly.Content);
            TestAssert.Equal(disabled.Annotations.Count, keyOnly.Annotations.Count);
        }
    }

    public static void PptxEmptyDataLabelKeysDoNotCreateBackgrounds()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        void Background(XElement labels) => labels.AddFirst(new XElement(C + "spPr",
            new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "F7E188")))));
        var disabled = RenderVisibilityCase(VisibilityCases[0].Input, key: false, value: false, Background);
        var keyOnly = RenderVisibilityCase(VisibilityCases[0].Input, key: true, value: false, Background);
        TestAssert.Equal(disabled.Content, keyOnly.Content);
        TestAssert.True(!PptxChartPaletteTests.ReadColoredPaints(keyOnly.Content)
            .Any(color => color.Red == 0xF7 && color.Green == 0xE1 && color.Blue == 0x88),
            "Empty label units do not paint their authored background.");
    }

    public static void PptxVisibleDataLabelsKeepKeysAcrossPlotKinds()
    {
        foreach (var source in VisibilityCases)
        {
            var plain = RenderVisibilityCase(source.Input, key: false, value: true);
            var keys = RenderVisibilityCase(source.Input, key: true, value: true);
            int before = PptxChartPaletteTests.ReadColoredPaints(plain.Content, strokes: source.Stroke).Count();
            int after = PptxChartPaletteTests.ReadColoredPaints(keys.Content, strokes: source.Stroke).Count();
            TestAssert.Equal(source.Keys, after - before);
        }
    }

    public static void PptxCustomDataLabelTextKeepsItsKeysWithoutEmptySiblings()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var page = RenderVisibilityCase(VisibilityCases[0].Input, key: true, value: false, labels =>
            labels.AddFirst(new XElement(C + "dLbl", new XElement(C + "idx", new XAttribute("val", "2")),
                new XElement(C + "tx", new XElement(C + "rich", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"),
                    new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", "1800"))),
                        new XElement(a + "r", new XElement(a + "t", "NamedPoint"))))))));
        // Three plot paths, three legend keys and three custom-label keys.
        TestAssert.Equal(9, PptxChartPaletteTests.ReadColoredPaints(page.Content, strokes: true).Count());
    }

    public static void PptxEmptyFormattedDataLabelsDoNotPaintKeys()
    {
        foreach (var (input, component) in new[] {
            ("BuiltinPalettes/line-3.pptx", "showPercent"),
            ("BuiltinPalettes/line-3.pptx", "showBubbleSize"),
            ("BuiltinPalettes/area-3.pptx", "showPercent"),
            ("BuiltinPalettes/area-3.pptx", "showBubbleSize"),
            ("BuiltinPalettes/count-series-3.pptx", "showPercent"),
            ("BuiltinPalettes/count-series-3.pptx", "showBubbleSize"),
            ("BuiltinPalettes/pie-vary.pptx", "showBubbleSize"),
            ("pptx-ladder-11-chart-scatter-clusters-port.pptx", "showPercent") })
        {
            var absent = RenderVisibilityCase(input, key: true, value: false);
            var empty = RenderVisibilityCase(input, key: true, value: false,
                labels => labels.Element(C + component)!.SetAttributeValue("val", "1"));
            foreach (bool strokes in new[] { false, true })
                TestAssert.True(PptxChartPaletteTests.ReadColoredPaints(absent.Content, strokes: strokes)
                    .SequenceEqual(PptxChartPaletteTests.ReadColoredPaints(empty.Content, strokes: strokes)),
                    input + ": a component with no formatted text cannot paint a standalone key.");
        }
    }

    private static PdfPage RenderVisibilityCase(string input, bool key, bool value, Action<XElement>? customize = null)
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", input));
        using var archive = ZipFile.OpenRead(path);
        var parts = new Dictionary<string, byte[]>();
        foreach (var entry in archive.Entries)
        {
            using var source = entry.Open();
            using var copy = new MemoryStream();
            source.CopyTo(copy);
            parts.Add(entry.FullName, copy.ToArray());
        }
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/charts/chart1.xml"]));
        var plot = xml.Descendants(C + "plotArea").Single().Elements().Single(element => element.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        var labels = plot.Element(C + "dLbls");
        if (labels is null)
        {
            labels = new XElement(C + "dLbls");
            plot.Elements(C + "ser").Last().AddAfterSelf(labels);
        }
        foreach (var component in Components)
        {
            var flag = labels.Element(C + component.Element);
            if (flag is null) { flag = new XElement(C + component.Element); labels.Add(flag); }
            flag.SetAttributeValue("val", component.Element == "showLegendKey" ? key ? "1" : "0" :
                component.Element == "showVal" && value ? "1" : "0");
        }
        customize?.Invoke(labels);
        parts["ppt/charts/chart1.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
        using var bytes = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(bytes, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default)
            .ToBlockingEnumerable().Single();
    }
}
