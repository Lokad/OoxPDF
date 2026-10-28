using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxBulletChoiceTests
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static void PptxBulletChoicesPreserveIndependentOfficeSourceSelection()
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("pptx-bullet-choices-office.json")));
        string source = Case("pptx-bullet-choices.pptx");
        TestAssert.Equal(gold.RootElement.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source))));
        WithPackage(source, (document, package, scene) =>
        {
            var resolver = new PresentationFontResolver(new PortableRenderingTests.PinnedResolver());
            int index = 0;
            foreach (var node in scene.Slides[0].SlideNodes)
            {
                string before = node.Source.ToString(SaveOptions.DisableFormatting);
                var spans = PptxRenderer.BuildSceneFedTextSpans(node, document, scene.Theme, scene.Slides[0].SlideColorMap,
                    1, false, [], resolver, default);
                var bullet = spans.Single(span => span.Run.Text == "*").Run;
                var expected = gold.RootElement.GetProperty("Rows")[index];
                TestAssert.Equal(expected.GetProperty("NominalSize").GetDouble(), bullet.FontSize);
                TestAssert.Equal(expected.GetProperty("FontFamily").GetString(), bullet.FontFamily);
                var color = expected.GetProperty("Color").GetInt32();
                TestAssert.Equal(new RgbColor((byte)(color >> 16), (byte)(color >> 8), (byte)color), bullet.Color);
                TestAssert.Equal(before, node.Source.ToString(SaveOptions.DisableFormatting));
                index++;
            }
            TestAssert.Equal(8, index);
            var paragraph = PptxRenderer.InspectTextFrameModels(document, package, 0)[5].Paragraphs.Single();
            TestAssert.Equal("FollowText", paragraph.BulletFontTypefaceSource);
            TestAssert.True(paragraph.BulletResolvedFontTypeface is null, "Follow-text font selection remains owned by the text run.");
        });
    }

    public static void PptxBulletChoicesSurviveTableAndPlaceholderInheritance()
    {
        var original = Parts();
        foreach (bool table in new[] { false, true })
        {
            var parts = new Dictionary<string, string>(original);
            var slide = XDocument.Parse(parts["ppt/slides/slide1.xml"]);
            var shapes = slide.Descendants(P + "sp").ToArray();
            var layout = new XElement(P + "sldLayout", new XAttribute(XNamespace.Xmlns + "p", P),
                new XAttribute(XNamespace.Xmlns + "a", A), new XElement(P + "cSld", new XElement(P + "spTree")));
            for (int i = 0; i < shapes.Length; i++)
            {
                XElement shape = shapes[i];
                if (table)
                {
                    var transform = shape.Element(P + "spPr")!.Element(A + "xfrm")!;
                    var text = new XElement(A + "txBody", shape.Element(P + "txBody")!.Elements().Select(element => new XElement(element)));
                    var cell = new XElement(A + "tc", text, new XElement(A + "tcPr"));
                    var row = new XElement(A + "tr", new XAttribute("h", 635000), cell);
                    var tableElement = new XElement(A + "tbl", new XElement(A + "tblPr"),
                        new XElement(A + "tblGrid", new XElement(A + "gridCol", new XAttribute("w", 7620000))), row);
                    var frame = new XElement(P + "graphicFrame", new XElement(P + "nvGraphicFramePr",
                        new XElement(P + "cNvPr", new XAttribute("id", i + 2), new XAttribute("name", "Cell " + i)), new XElement(P + "cNvGraphicFramePr"), new XElement(P + "nvPr")),
                        new XElement(P + "xfrm", transform.Elements().Select(element => new XElement(element))),
                        new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", A.NamespaceName.Replace("/main", "/table")),
                            tableElement)));
                    shape.ReplaceWith(frame);
                }
                else
                {
                    shape.Element(P + "nvSpPr")!.Element(P + "nvPr")!.Add(new XElement(P + "ph", new XAttribute("idx", i + 1)));
                    var inherited = new XElement(shape);
                    foreach (var properties in inherited.Descendants(A + "pPr")) properties.RemoveNodes();
                    layout.Element(P + "cSld")!.Element(P + "spTree")!.Add(inherited);
                    shape.Element(P + "txBody")!.Element(A + "lstStyle")!.RemoveNodes();
                }
            }
            parts["ppt/slides/slide1.xml"] = slide.ToString();
            if (!table)
            {
                var relationship = XDocument.Parse(parts["ppt/slides/_rels/slide1.xml.rels"]).Root!.Elements()
                    .Single(element => ((string?)element.Attribute("Type"))?.EndsWith("/slideLayout", StringComparison.Ordinal) == true);
                string target = OoxPath.ResolveRelationshipTarget("/ppt/slides/slide1.xml", (string)relationship.Attribute("Target")!);
                parts[target.TrimStart('/')] = layout.ToString();
            }
            string path = TestFixtures.WriteTempPackage(".pptx", parts);
            WithPackage(path, (document, package, _) =>
            {
                var frames = table ? PptxRenderer.InspectTableTextFrameModels(document, package, 0)
                    : PptxRenderer.InspectTextFrameModels(document, package, 0);
                TestAssert.Equal(8, frames.Count);
                var paragraphs = frames.Select(frame => frame.Paragraphs.Single()).ToArray();
                TestAssert.Equal("Points", paragraphs[2].BulletSizeKind);
                TestAssert.Equal("2400", paragraphs[2].BulletSizeValue);
                TestAssert.Equal("Text", paragraphs[4].BulletSizeKind);
                TestAssert.Equal("FollowText", paragraphs[5].BulletFontTypefaceSource);
                TestAssert.True(paragraphs[6].BulletColor is null, "Follow-text color must not inherit red.");
                TestAssert.Equal("FF0000", paragraphs[7].BulletColor);
                TestAssert.Equal("Arial", paragraphs[7].BulletResolvedFontTypeface);
            });
        }
    }

    public static void PptxBulletChoicesKeepMarkerKindsAndIgnoreLateAuthoredProperties()
    {
        foreach (string direct in new[] { "<a:buNone/>", "<a:buAutoNum type=\"arabicPeriod\"/>", "<a:buChar char=\"*\"/><a:buSzPts val=\"4800\"/>" })
        {
            var parts = Parts();
            var slide = XDocument.Parse(parts["ppt/slides/slide1.xml"]);
            var shape = slide.Descendants(P + "sp").First();
            foreach (var other in slide.Descendants(P + "sp").Skip(1).ToArray()) other.Remove();
            var style = shape.Descendants(A + "lvl1pPr").Single();
            style.ReplaceNodes(XElement.Parse("<a:buSzPct xmlns:a=\"" + A + "\" val=\"50000\"/>"),
                new XElement(A + "buFont", new XAttribute("typeface", "Arial")), new XElement(A + "buChar", new XAttribute("char", "*")));
            var properties = XElement.Parse("<a:pPr xmlns:a=\"" + A + "\">" + direct + "</a:pPr>");
            shape.Descendants(A + "pPr").Single().ReplaceWith(properties);
            parts["ppt/slides/slide1.xml"] = slide.ToString();
            WithPackage(TestFixtures.WriteTempPackage(".pptx", parts), (document, package, scene) =>
            {
                var paragraph = PptxRenderer.InspectTextFrameModels(document, package, 0).Single().Paragraphs.Single();
                if (direct.Contains("buNone", StringComparison.Ordinal)) TestAssert.Equal("None", paragraph.BulletKind);
                else
                {
                    TestAssert.Equal(direct.Contains("buAutoNum", StringComparison.Ordinal) ? "AutoNumber" : "Character", paragraph.BulletKind);
                    TestAssert.Equal("Percent", paragraph.BulletSizeKind);
                    TestAssert.Equal("50000", paragraph.BulletSizeValue);
                    var resolver = new PresentationFontResolver(new PortableRenderingTests.PinnedResolver());
                    var spans = PptxSceneTextAgreementTests.ReadXmlSpans(scene.Slides[0].SlideNodes.Single(), document, scene, resolver: resolver);
                    TestAssert.Equal(12d, spans.First().Run.FontSize);
                }
            });
        }
    }

    private static Dictionary<string, string> Parts()
    {
        using var zip = ZipFile.OpenRead(Case("pptx-bullet-choices.pptx"));
        return zip.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal) || entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, entry => { using var reader = new StreamReader(entry.Open()); return reader.ReadToEnd(); });
    }
    private static void WithPackage(string path, Action<PptxDocument, OoxPackage, PptxScene> check)
    {
        using var input = File.OpenRead(path);
        var package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        check(document, package, new PptxSceneBuilder().Build(document, package, default));
    }
    private static string Case(string file) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes", file));
}
