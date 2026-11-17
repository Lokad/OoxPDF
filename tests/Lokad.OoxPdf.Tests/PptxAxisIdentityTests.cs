using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAxisIdentityTests
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public static void PptxChartAxisNumericSpellingsPreserveScaleAndGeometry()
    {
        string cases = Cases();
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office-labels.json")));
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (string name in new[] { "solo-column", "shared-reversed", "distinct-secondary" })
        {
            var row = oracle.RootElement.GetProperty("Rows").EnumerateArray()
                .Single(row => row.GetProperty("Id").GetString() == name);
            byte[] source = File.ReadAllBytes(Path.Combine(cases, row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(source)));
            PdfPage baseline = Render(source, _ => { }, resolver).Page;
            string[] expected = row.GetProperty("OfficeNumericLabels").EnumerateArray()
                .Select(value => value.GetString()!).Order(StringComparer.Ordinal).ToArray();
            TestAssert.True(expected.SequenceEqual(NumericLabels(baseline)), "Independent Office scale: " + name);

            foreach (string owner in new[] { "value-definition", "category-definition", "plot-reference", "cross-reference", "all" })
            {
                var actual = Render(source, xml =>
                {
                    foreach (XElement id in xml.Descendants().Where(element => element.Name == C + "axId" || element.Name == C + "crossAx"))
                    {
                        string parent = id.Parent!.Name.LocalName;
                        bool selected = owner == "all" || (owner switch
                        {
                            "value-definition" => parent == "valAx" && id.Name == C + "axId",
                            "category-definition" => parent == "catAx" && id.Name == C + "axId",
                            "plot-reference" => parent.EndsWith("Chart", StringComparison.Ordinal),
                            _ => id.Name == C + "crossAx"
                        });
                        if (selected) id.SetAttributeValue("val", " +00" + (string)id.Attribute("val")! + " ");
                    }
                }, resolver).Page;
                TestAssert.True(expected.SequenceEqual(NumericLabels(actual)), "Numeric identity retains authored scale: " + name + "/" + owner);
                TestAssert.Equal(baseline.Content, actual.Content);
            }
        }
    }

    public static void PptxChartAxisIdentityPreservesRawAndMalformedRecovery()
    {
        byte[] source = File.ReadAllBytes(Path.Combine(Cases(), "solo-column.pptx"));
        var resolver = new PortableRenderingTests.PinnedResolver();
        PdfPage baseline = Render(source, _ => { }, resolver).Page;
        foreach (var pair in new[]
        {
            (Definition: " +004294967295 ", Reference: "4294967295", Match: true),
            (Definition: "-0", Reference: "+000", Match: true),
            (Definition: "opaque", Reference: "opaque", Match: true),
            (Definition: "Opaque", Reference: "opaque", Match: false),
            (Definition: "04294967296", Reference: "4294967296", Match: false),
            (Definition: "-1", Reference: "4294967295", Match: false),
            (Definition: "\u00a02140495176\u00a0", Reference: "2140495176", Match: false),
            (Definition: "", Reference: "2140495176", Match: false)
        })
        {
            var actual = Render(source, xml =>
            {
                string oldId = (string)xml.Descendants(C + "valAx").Single().Element(C + "axId")!.Attribute("val")!;
                foreach (var id in xml.Descendants().Where(element => element.Name == C + "axId" || element.Name == C + "crossAx"))
                    if ((string?)id.Attribute("val") == oldId)
                        id.SetAttributeValue("val", id.Parent?.Name == C + "valAx" && id.Name == C + "axId" ? pair.Definition : pair.Reference);
            }, resolver, inspect: true);
            TestAssert.Equal(pair.Match, baseline.Content == actual.Page.Content);
            if (pair.Definition.Length != 0)
                TestAssert.Equal(pair.Definition, actual.Scene!.Slides.Single().SlideNodes.Single().Chart!.Axes
                    .Single(axis => axis.AxisKind == PptxSceneChartAxisKind.Value).Id);
        }
    }

    private static string Cases() => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ComposedAxes"));

    private static string[] NumericLabels(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
        usePageCoordinates: false, fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal))
        .Select(show => show.Text).Order(StringComparer.Ordinal).ToArray();

    private static (PdfPage Page, PptxScene? Scene) Render(byte[] source, Action<XDocument> change,
        PortableRenderingTests.PinnedResolver resolver, bool inspect = false)
    {
        using var stream = new MemoryStream();
        stream.Write(source); stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            change(xml);
            entry.Delete();
            using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
            xml.Save(output);
        }
        stream.Position = 0;
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        var scene = inspect ? new PptxSceneBuilder().Build(document, package, default) : null;
        var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
        var page = new PptxRenderer(resolver).RenderPagesAsync(document, package, diagnostics.Add, default)
            .ToBlockingEnumerable().Single();
        TestAssert.Equal(0, diagnostics.Count);
        return (page, scene);
    }
}
