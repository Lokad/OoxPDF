using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxMasterThemeTests
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static void PptxMasterThemesPreserveOfficeFillOwnership()
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("pptx-master-themes-office.json")));
        foreach (var expected in gold.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("OfficeApp").GetString() == "PowerPoint"))
        {
            string path = Case(expected.GetProperty("File").GetString()!);
            TestAssert.Equal(expected.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            WithDeck(Parts(path), (document, package) =>
            {
                var pages = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPages(document, package, null, default).ToArray();
                TestAssert.Equal(2, pages.Length);
                // Exercise paint before scene metadata, so the preceding runtime
                // also fails this regression on the demonstrated wrong color.
                for (int index = 0; index < pages.Length; index++)
                {
                    var rgb = expected.GetProperty("Pages")[index].GetProperty("ThemeRectangles")[0].GetProperty("Fill");
                    string fill = rgb[2].GetDouble() == 1d ? "0 0 1 rg" : "1 0 0 rg";
                    TestAssert.Contains(fill, pages[index].Content);
                }
                var scene = new PptxSceneBuilder().Build(document, package, default);
                for (int index = 0; index < scene.Slides.Count; index++)
                {
                    bool blue = expected.GetProperty("Pages")[index].GetProperty("ThemeRectangles")[0].GetProperty("Fill")[2].GetDouble() == 1d;
                    TestAssert.Equal(blue ? new RgbColor(0, 0, 255) : new RgbColor(255, 0, 0), scene.Slides[index].SlideNodes.Last().Shape!.Fill.Color);
                }
            });
        }
    }

    public static void PptxMasterThemesCarryBackgroundAndTextOwnership()
    {
        var parts = Parts();
        for (int index = 1; index <= 2; index++)
        {
            Edit(parts, $"ppt/theme/theme{index}.xml", xml => xml.Descendants(A + "minorFont").Single().Element(A + "latin")!
                .SetAttributeValue("typeface", index == 1 ? "Arial" : "Courier New"));
            Edit(parts, $"ppt/slideMasters/slideMaster{index}.xml", xml => xml.Root!.Element(P + "cSld")!.AddFirst(
                new XElement(P + "bg", new XElement(P + "bgPr", new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "accent1")))))));
            Edit(parts, $"ppt/slides/slide{index}.xml", xml => xml.Descendants(A + "latin").Single().SetAttributeValue("typeface", "+mn-lt"));
        }
        WithDeck(parts, (document, package) =>
        {
            var scene = new PptxSceneBuilder().Build(document, package, default);
            TestAssert.Equal(new RgbColor(255, 0, 0), scene.Slides[0].MasterBackground.Color);
            TestAssert.Equal(new RgbColor(0, 0, 255), scene.Slides[1].MasterBackground.Color);
            for (int index = 0; index < 2; index++)
                TestAssert.Equal(index == 0 ? "Arial" : "Courier New", PptxRenderer.InspectTextFrameModels(document, package, index).First().Paragraphs.Single().Runs.Single().Typeface);
            var pages = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPages(document, package, null, default).ToArray();
            TestAssert.Contains("1 0 0 rg", pages[0].Content);
            TestAssert.Contains("0 0 1 rg", pages[1].Content);
        });
    }

    public static void PptxMasterThemesShareActualPartsAndKeepFallbacks()
    {
        foreach (bool sameMaster in new[] { true, false })
        {
            var parts = Parts();
            Edit(parts, sameMaster ? "ppt/slides/_rels/slide2.xml.rels" : "ppt/slideMasters/_rels/slideMaster2.xml.rels", xml =>
                xml.Root!.Elements().Single(e => ((string?)e.Attribute("Type"))?.EndsWith(sameMaster ? "/slideLayout" : "/theme", StringComparison.Ordinal) == true)
                    .SetAttributeValue("Target", sameMaster ? "../slideLayouts/slideLayout1.xml" : "../theme/theme1.xml"));
            WithDeck(parts, (document, package) =>
            {
                var scene = new PptxSceneBuilder().Build(document, package, default);
                TestAssert.True(ReferenceEquals(scene.Slides[0].Theme, scene.Slides[1].Theme), "One parsed model per actual theme part, including different masters sharing it.");
            });
        }
        foreach (bool noTheme in new[] { false, true })
        {
            var parts = Parts();
            Edit(parts, "ppt/slideMasters/_rels/slideMaster2.xml.rels", xml => xml.Root!.Elements().Where(IsTheme).Remove());
            if (noTheme) Edit(parts, "ppt/slideMasters/_rels/slideMaster1.xml.rels", xml => xml.Root!.Elements().Where(IsTheme).Remove());
            WithDeck(parts, (document, package) =>
            {
                var scene = new PptxSceneBuilder().Build(document, package, default);
                TestAssert.True(ReferenceEquals(scene.Slides[0].Theme, scene.Slides[1].Theme), "Missing master theme retains the presentation/first-master fallback.");
                TestAssert.Equal(noTheme, ReferenceEquals(PptxTheme.Empty, scene.Slides[1].Theme));
                using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => new PptxSceneBuilder().Build(document, package, cancelled.Token));
            });
        }
    }

    private static bool IsTheme(XElement element) => ((string?)element.Attribute("Type"))?.EndsWith("/theme", StringComparison.Ordinal) == true;
    private static Dictionary<string, byte[]> Parts(string? path = null)
    {
        using var zip = ZipFile.OpenRead(path ?? Case("pptx-master-themes-distinct.pptx"));
        return zip.Entries.ToDictionary(entry => entry.FullName, entry => { using var data = new MemoryStream(); using var stream = entry.Open(); stream.CopyTo(data); return data.ToArray(); });
    }
    private static void Edit(Dictionary<string, byte[]> parts, string name, Action<XDocument> edit)
    {
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[name]));
        edit(xml);
        parts[name] = Encoding.UTF8.GetBytes(xml.ToString());
    }
    private static void WithDeck(Dictionary<string, byte[]> parts, Action<PptxDocument, OoxPackage> check)
    {
        using var stream = TestFixtures.CreateZipPackage(parts);
        var package = OoxPackage.Open(stream, default);
        check(new PptxReader().Read(package, default), package);
    }
    private static string Case(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name));
}
