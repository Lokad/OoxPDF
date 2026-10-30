using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxInheritanceTests
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void PptxInheritanceOfficeBackgroundOverridesMaster()
    {
        WithDeck(Parts("background-inheritance"), (document, package) =>
        {
            var pages = Render(document, package);
            TestAssert.Equal(2, pages.Length);
            foreach (var page in pages) CheckBackground(page, "0 0 1 rg");
            TestAssert.True(pages[0].ExtGStates.Any(state => state.FillAlpha == 0.5d),
                "Keep authored transparency; the overridden red master must not be underneath it.");
        });
    }

    public static void PptxInheritanceBackgroundFallbacksPaintOneFill()
    {
        foreach (int owner in new[] { -1, 0, 1, 2 })
        foreach (bool transparent in new[] { false, true })
        {
            var parts = Parts("background-inheritance");
            string[] colors = ["FF0000", "00FF00", "0000FF"];
            string[] names = ["ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/slides/slide1.xml"];
            for (int index = 0; index < names.Length; index++)
            {
                int source = index;
                Edit(parts, names[index], xml =>
                {
                    var story = xml.Root!.Element(P + "cSld")!;
                    story.Elements(P + "bg").Remove();
                    if (source > owner) return;
                    var color = new XElement(A + "srgbClr", new XAttribute("val", colors[source]));
                    if (transparent) color.Add(new XElement(A + "alpha", new XAttribute("val", "50000")));
                    story.AddFirst(new XElement(P + "bg", new XElement(P + "bgPr", new XElement(A + "solidFill", color))));
                });
            }
            WithDeck(parts, (document, package) => CheckBackground(Render(document, package)[0],
                owner switch { 0 => "1 0 0 rg", 1 => "0 1 0 rg", 2 => "0 0 1 rg", _ => "1 g" }));
        }
    }

    public static void PptxInheritanceOfficeLayoutAndSlideMasterPermissions()
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("pptx-inheritance-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("Id").GetString() != "background-inheritance"))
        {
            WithDeck(Parts(row.GetProperty("Id").GetString()!), (document, package) =>
            {
                var pages = Render(document, package);
                var scene = new PptxSceneBuilder().Build(document, package, default);
                for (int index = 0; index < pages.Length; index++)
                {
                    int expected = row.GetProperty("MasterRectangles")[index].GetInt32();
                    TestAssert.Equal(expected, Regex.Matches(pages[index].Content, @"(?m)^1 0 0 rg\r?$").Count);
                    TestAssert.Equal(expected == 1, scene.Slides[index].ShowMasterShapes);
                }
            });
        }
        // Missing layouts retain the slide's ordinary default permission.
        var missing = Parts("layout-master-suppression");
        Edit(missing, "ppt/slides/_rels/slide1.xml.rels", xml =>
            xml.Root!.Elements().Where(e => ((string?)e.Attribute("Type"))?.EndsWith("/slideLayout", StringComparison.Ordinal) == true).Remove());
        WithDeck(missing, (document, package) =>
            TestAssert.True(new PptxSceneBuilder().Build(document, package, default).Slides[0].ShowMasterShapes, "Absent layout does not suppress artwork."));
    }

    public static void PptxInheritanceLayoutPermissionReachesGroupsLinksAndDiagnostics()
    {
        foreach (bool shown in new[] { false, true })
        {
            var parts = Parts("layout-master-suppression");
            Edit(parts, "ppt/slides/slide1.xml", xml => xml.Root!.SetAttributeValue("showMasterSp", "1"));
            Edit(parts, "ppt/slideLayouts/slideLayout1.xml", xml => xml.Root!.SetAttributeValue("showMasterSp", shown ? "1" : "0"));
            Edit(parts, "ppt/slideMasters/slideMaster1.xml", xml =>
            {
                var root = xml.Root!;
                root.SetAttributeValue(XNamespace.Xmlns + "review", "urn:public-inheritance-review");
                root.SetAttributeValue(XName.Get("MustUnderstand", "http://schemas.openxmlformats.org/markup-compatibility/2006"), "review");
                var tree = root.Element(P + "cSld")!.Element(P + "spTree")!;
                var shape = tree.Elements(P + "sp").Single();
                shape.Element(P + "nvSpPr")!.Element(P + "cNvPr")!.Add(new XElement(A + "hlinkClick", new XAttribute(R + "id", "reviewLink")));
                shape.ReplaceWith(new XElement(P + "grpSp", new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr", new XAttribute("id", "40"), new XAttribute("name", "Master group")), new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")), new XElement(P + "grpSpPr"), new XElement(shape)));
            });
            Edit(parts, "ppt/slideMasters/_rels/slideMaster1.xml.rels", xml => xml.Root!.Add(
                new XElement(xml.Root.Name.Namespace + "Relationship", new XAttribute("Id", "reviewLink"), new XAttribute("Type", R.NamespaceName + "/hyperlink"), new XAttribute("Target", "https://example.test/master"), new XAttribute("TargetMode", "External"))));
            WithDeck(parts, (document, package) =>
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = Render(document, package, diagnostics)[0];
                TestAssert.Equal(shown ? 1 : 0, page.Annotations.Count);
                TestAssert.Equal(shown, page.Content.Contains("1 0 0 rg", StringComparison.Ordinal));
                TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_MUST_UNDERSTAND" && d.PartName == "/ppt/slideMasters/slideMaster1.xml"),
                    "Master part compatibility checks survive artwork suppression.");
            });
            Edit(parts, "ppt/slideMasters/slideMaster1.xml", xml => xml.Descendants(A + "prstGeom").Single().SetAttributeValue("prst", "wedgeRoundRectCallout"));
            WithDeck(parts, (document, package) =>
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                Render(document, package, diagnostics);
                TestAssert.Equal(shown, diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_CALLOUT" && d.SlideIndex == 1));
            });
        }
    }

    private static void CheckBackground(PdfPage page, string expected)
    {
        TestAssert.Equal(expected, Regex.Match(page.Content, @"(?m)^(?:[\d.]+ [\d.]+ [\d.]+ rg|[\d.]+ g)\r?$").Value.TrimEnd('\r'));
        string rectangle = "0 0 " + PdfDocumentWriter.FormatNumber(page.Width) + " " + PdfDocumentWriter.FormatNumber(page.Height) + " re f*";
        TestAssert.Equal(1, page.Content.Split('\n').Count(line => line.TrimEnd('\r') == rectangle));
    }
    private static PdfPage[] Render(PptxDocument document, OoxPackage package, List<OoxPdfDiagnostic>? diagnostics = null) =>
        new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, diagnostics is null ? null : diagnostics.Add, default).ToBlockingEnumerable().ToArray();
    private static Dictionary<string, byte[]> Parts(string id)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("pptx-inheritance-office.json")));
        var row = gold.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == id);
        string path = Case(row.GetProperty("File").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.ToDictionary(entry => entry.FullName, entry => { using var buffer = new MemoryStream(); using var input = entry.Open(); input.CopyTo(buffer); return buffer.ToArray(); });
    }
    private static void Edit(Dictionary<string, byte[]> parts, string name, Action<XDocument> edit)
    {
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[name])); edit(xml);
        parts[name] = Encoding.UTF8.GetBytes(xml.ToString());
    }
    private static void WithDeck(Dictionary<string, byte[]> parts, Action<PptxDocument, OoxPackage> check)
    {
        using var input = TestFixtures.CreateZipPackage(parts);
        var package = OoxPackage.Open(input, default);
        check(new PptxReader().Read(package, default), package);
    }
    private static string Case(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name));
}
