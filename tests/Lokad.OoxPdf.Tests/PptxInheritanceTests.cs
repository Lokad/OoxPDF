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

internal static partial class PptxInheritanceTests
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly (string Part, string Color, string Paint)[] BackgroundLayers =
    [
        ("ppt/slideMasters/slideMaster1.xml", "FF0000", "1 0 0 rg"),
        ("ppt/slideLayouts/slideLayout1.xml", "00FF00", "0 1 0 rg"),
        ("ppt/slides/slide1.xml", "0000FF", "0 0 1 rg")
    ];

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
            for (int index = 0; index < BackgroundLayers.Length; index++)
            {
                SetBackground(parts, BackgroundLayers[index].Part, index <= owner ? SolidBackground(index, transparent) : null);
            }
            WithDeck(parts, (document, package) => CheckBackground(Render(document, package)[0],
                owner < 0 ? "1 g" : BackgroundLayers[owner].Paint));
        }
    }

    public static void PptxInheritanceResolvesSolidBackgroundStyleReferences()
    {
        (int Owner, string Index, bool Overridden, string Paint, string? Warning, double? Alpha)[] cases =
        [
            (0, "1001", false, "1 0 1 rg", null, null),
            (1, "1", false, "0 1 1 rg", null, null),
            (2, "1002", false, "1 0 1 rg", null, .25d),
            (2, "2", false, "0 g", null, null),
            (2, "0", false, "0 1 0 rg", null, null),
            (2, "1000", false, "0 1 0 rg", null, null),
            (2, "3", false, "0 1 0 rg", null, null),
            (0, "invalid", false, "1 g", BackgroundLayers[0].Part, null),
            (1, "1003", false, "1 0 0 rg", BackgroundLayers[1].Part, null),
            (2, "1004", false, "0 1 0 rg", BackgroundLayers[2].Part, null),
            (1, "1003", true, "0 0 1 rg", null, null)
        ];
        var template = Parts("background-inheritance");
        Edit(template, "ppt/theme/theme1.xml", xml =>
        {
            var scheme = xml.Descendants(A + "clrScheme").Single();
            scheme.Element(A + "accent1")!.ReplaceNodes(new XElement(A + "srgbClr", new XAttribute("val", "FF00FF")));
            scheme.Element(A + "accent2")!.ReplaceNodes(new XElement(A + "srgbClr", new XAttribute("val", "00FFFF")));
            var format = xml.Descendants(A + "fmtScheme").Single();
            format.Element(A + "fillStyleLst")!.ReplaceNodes(PlaceholderFill(),
                new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", "000000"))),
                new XElement(A + "noFill"));
            var backgrounds = format.Element(A + "bgFillStyleLst")!;
            var gradient = new XElement(backgrounds.Element(A + "gradFill")!);
            gradient.Element(A + "gsLst")!.Elements(A + "gs").First().Element(A + "schemeClr")!.Add(HalfAlpha());
            backgrounds.ReplaceNodes(PlaceholderFill(), PlaceholderFill(withAlpha: true), gradient);
        });
        foreach (var (owner, styleIndex, overridden, paint, warning, alpha) in cases)
        {
            var parts = new Dictionary<string, byte[]>(template);
            for (int index = 0; index < BackgroundLayers.Length; index++)
            {
                XElement? background = index < owner || (overridden && index == 2) ? SolidBackground(index) : null;
                if (index == owner)
                {
                    var color = new XElement(A + "schemeClr", new XAttribute("val", "accent1"));
                    if (alpha.HasValue) color.Add(HalfAlpha());
                    background = new XElement(P + "bgRef", new XAttribute("idx", styleIndex), color);
                }
                SetBackground(parts, BackgroundLayers[index].Part, background);
            }
            if (styleIndex == "1")
            {
                var masterMap = XDocument.Parse(Encoding.UTF8.GetString(parts[BackgroundLayers[0].Part])).Root!
                    .Element(P + "clrMap")!;
                for (int index = 1; index <= 2; index++)
                {
                    var map = new XElement(A + "overrideClrMapping", masterMap.Attributes());
                    map.SetAttributeValue("accent1", index == 1 ? "accent2" : "accent3");
                    Edit(parts, BackgroundLayers[index].Part, xml =>
                    {
                        xml.Root!.Elements(P + "clrMapOvr").Remove();
                        xml.Root.Element(P + "cSld")!.AddAfterSelf(new XElement(P + "clrMapOvr", map));
                    });
                }
            }
            WithDeck(parts, (document, package) =>
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = Render(document, package, diagnostics)[0];
                CheckBackground(page, paint);
                TestAssert.Equal(0, page.Shadings.Count);
                var diagnostic = diagnostics.SingleOrDefault(item => item.SlideIndex == 1);
                TestAssert.Equal(warning is null ? null : "/" + warning, diagnostic?.PartName);
                TestAssert.Equal(warning is null ? null : "PPTX_UNSUPPORTED_BACKGROUND_STYLE", diagnostic?.Id);
                if (alpha is { } expected)
                    TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == expected), "Reference and style opacity both apply.");
            });
        }

        static XElement PlaceholderFill(bool withAlpha = false) =>
            new(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr"), withAlpha ? HalfAlpha() : null));
        static XElement HalfAlpha() => new(A + "alpha", new XAttribute("val", "50000"));
    }

    public static void PptxInheritanceLinearBackgroundsPaintOneShading()
    {
        var template = Parts("background-inheritance");
        foreach (var (owner, fallback, overridden) in new[]
        {
            (0, "", false), (1, "", false), (2, "", false),
            (0, "mixed-alpha", false), (2, "mixed-alpha", false), (2, "title", false),
            (0, "invalid-color", false), (1, "invalid-color", false), (2, "invalid-color", false),
            (1, "invalid-color", true)
        })
        foreach (bool referenced in fallback.Length == 0 ? new[] { false, true } : new[] { false })
        {
            var parts = new Dictionary<string, byte[]>(template);
            for (int index = 0; index < BackgroundLayers.Length; index++)
            {
                XElement? background = index < owner || (overridden && index == 2) ? SolidBackground(index) : null;
                if (index == owner)
                {
                    var gradient = new XElement(A + "gradFill",
                        new XElement(A + "gsLst", Stop(0, "FF0000", 50000),
                            Stop(100000, fallback == "invalid-color" ? "NOTRGB" : "0000FF",
                                fallback == "mixed-alpha" ? 100000 : 50000)),
                        new XElement(A + "lin", new XAttribute("ang", 5400000)));
                    if (referenced)
                    {
                        XElement firstStop = gradient.Element(A + "gsLst")!.Elements(A + "gs").First();
                        XElement referenceColor = firstStop.Elements().Single();
                        referenceColor.ReplaceWith(new XElement(A + "schemeClr", new XAttribute("val", "phClr")));
                        Edit(parts, "ppt/theme/theme1.xml", xml =>
                            xml.Descendants(A + "bgFillStyleLst").Single().Elements(A + "gradFill").Single().ReplaceWith(gradient));
                        background = new XElement(P + "bgRef", new XAttribute("idx", "1003"), referenceColor);
                    }
                    else
                    {
                        background = new XElement(P + "bgPr", gradient);
                        if (fallback == "title") background.SetAttributeValue("shadeToTitle", "1");
                    }
                }
                SetBackground(parts, BackgroundLayers[index].Part, background);
            }

            WithDeck(parts, (document, package) =>
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = Render(document, package, diagnostics)[0];
                TestAssert.True(!diagnostics.Any(d => d.SlideIndex == 1 && d.Id == "PPTX_UNSUPPORTED_BACKGROUND_STYLE"),
                    "A supported theme gradient paints without a background style warning.");
                var gradientWarning = diagnostics.SingleOrDefault(d => d.SlideIndex == 1 && d.Id == "PPTX_UNSUPPORTED_GRADIENT_FILL");
                string? expectedOwner = fallback.Length != 0 && !overridden ? "/" + BackgroundLayers[owner].Part : null;
                TestAssert.Equal(expectedOwner, gradientWarning?.PartName);
                if (fallback.Length != 0)
                {
                    TestAssert.Equal(0, page.Shadings.Count);
                    int paintedOwner = overridden ? 2 : owner - 1;
                    CheckBackground(page, paintedOwner < 0 ? "1 g" : BackgroundLayers[paintedOwner].Paint);
                    return;
                }

                PdfShadingResource resource = page.Shadings.Single();
                var fill = (PdfAxialShading)resource.Shading;
                TestAssert.True(Math.Abs(fill.X0 - page.Width / 2d) < 1e-8 &&
                    Math.Abs(fill.X1 - page.Width / 2d) < 1e-8 &&
                    Math.Abs(fill.Y0 - page.Height) < 1e-8 && Math.Abs(fill.Y1) < 1e-8,
                    "The background gradient uses the full slide frame.");
                TestAssert.Equal(new PdfShadingStop(0d, 255, 0, 0), fill.Stops.First());
                TestAssert.Equal(new PdfShadingStop(1d, 0, 0, 255), fill.Stops.Last());
                TestAssert.Contains("/" + resource.ResourceName + " sh", page.Content);
                TestAssert.DoesNotContain(FullSlideFill(page), page.Content);
                TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == .5d), "The selected background retains its uniform alpha.");
            });
        }

        static XElement Stop(int offset, string rgb, int alpha) =>
            new(A + "gs", new XAttribute("pos", offset),
                new XElement(A + "srgbClr", new XAttribute("val", rgb), new XElement(A + "alpha", new XAttribute("val", alpha))));
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
        string rectangle = FullSlideFill(page);
        TestAssert.Equal(1, page.Content.Split('\n').Count(line => line.TrimEnd('\r') == rectangle));
    }

    private static string FullSlideFill(PdfPage page) =>
        "0 0 " + PdfDocumentWriter.FormatNumber(page.Width) + " " + PdfDocumentWriter.FormatNumber(page.Height) + " re f*";

    private static XElement SolidBackground(int owner, bool transparent = false)
    {
        var color = new XElement(A + "srgbClr", new XAttribute("val", BackgroundLayers[owner].Color));
        if (transparent) color.Add(new XElement(A + "alpha", new XAttribute("val", "50000")));
        return new XElement(P + "bgPr", new XElement(A + "solidFill", color));
    }

    private static void SetBackground(Dictionary<string, byte[]> parts, string part, XElement? background) =>
        Edit(parts, part, xml =>
        {
            var story = xml.Root!.Element(P + "cSld")!;
            story.Elements(P + "bg").Remove();
            if (background is not null) story.AddFirst(new XElement(P + "bg", background));
        });
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
        using var package = OoxPackage.Open(input, default);
        check(new PptxReader().Read(package, default), package);
    }
    private static string Case(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name));
}
