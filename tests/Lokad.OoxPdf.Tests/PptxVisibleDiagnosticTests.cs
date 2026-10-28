using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static class PptxVisibleDiagnosticTests
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageR = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static void PptxVisibleDiagnosticsExcludeHiddenMediaAndGraphicOwners()
    {
        byte[] baseline = Convert(Parts(null, "0", false)).Pdf;
        foreach (string kind in new[] { "video", "audio", "ole", "graphic" })
        foreach (string hidden in new[] { "1", "true", "0", "false" })
        foreach (bool group in new[] { false, true })
        {
            var result = Convert(Parts(kind, hidden, group));
            if (hidden is "1" or "true")
            {
                TestAssert.Equal(0, result.Diagnostics.Count);
                TestAssert.True(baseline.AsSpan().SequenceEqual(result.Pdf), "Hidden drawing content must not change the admitted PDF or diagnostics.");
            }
            else
            {
                string code = kind switch {
                    "video" => "PPTX_UNSUPPORTED_VIDEO", "audio" => "PPTX_UNSUPPORTED_AUDIO",
                    "ole" => "PPTX_UNSUPPORTED_OLE_OBJECT", _ => "PPTX_UNSUPPORTED_GRAPHIC_FRAME" };
                TestAssert.True(result.Diagnostics.Any(item => item.Id == code && item.PartName == "/ppt/slides/slide1.xml"),
                    "Visible unsupported content keeps its source-owned warning.");
                if (kind is "video" or "audio") TestAssert.True(!baseline.AsSpan().SequenceEqual(result.Pdf), "Visible media keeps its static poster.");
            }
        }
    }

    public static void PptxVisibleDiagnosticsKeepBackgroundOrphanAndGlobalWarnings()
    {
        var parts = Parts("video", "true", true);
        var source = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        var root = source.Root!;
        root.Element(P + "cSld")!.Element(P + "spTree")!.Add(new XElement(P + "oleObj"));
        root.Add(new XElement(P + "transition"), new XElement(P + "timing"));
        root.Element(P + "cSld")!.AddFirst(new XElement(P + "bg", new XElement(P + "bgPr", new XElement(A + "pattFill", new XAttribute("prst", "pct5")))));
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        root.Add(new XAttribute(XNamespace.Xmlns + "x", "urn:public-diagnostic-probe"), new XAttribute(XNamespace.Xmlns + "mc", mc),
            new XAttribute(mc + "MustUnderstand", "x"));
        parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(source.ToString());
        var result = Convert(parts);
        foreach (string code in new[] { "PPTX_UNSUPPORTED_OLE_OBJECT", "PPTX_UNSUPPORTED_TRANSITION", "PPTX_UNSUPPORTED_ANIMATION", "PPTX_UNSUPPORTED_PATTERN_FILL", "OOXML_MUST_UNDERSTAND" })
            TestAssert.True(result.Diagnostics.Any(item => item.Id == code), "Required non-scene/global warning remains: " + code);
        TestAssert.True(result.Diagnostics.All(item => item.Id != "PPTX_UNSUPPORTED_VIDEO"), "The unrelated hidden video remains excluded.");
    }

    public static void PptxVisibleDiagnosticsCancelBeforeAtomicPublication()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", Parts("video", "false", false));
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        byte[] original = [1, 2, 3, 4];
        File.WriteAllBytes(output, original);
        using var cancellation = new CancellationTokenSource();
        bool observed = false;
        try
        {
            TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(input, output,
                new OoxPdfOptions { FontResolver = new PortableRenderingTests.PinnedResolver(), DiagnosticSink = item =>
                {
                    if (item.Id == "PPTX_UNSUPPORTED_VIDEO") { observed = true; cancellation.Cancel(); }
                } }, cancellation.Token));
            TestAssert.True(observed, "Cancellation follows a genuinely visible unsupported object.");
            TestAssert.True(original.AsSpan().SequenceEqual(File.ReadAllBytes(output)), "Cancelled conversion must preserve existing output.");
            TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
        }
        finally { File.Delete(input); File.Delete(output); }
    }

    private static (byte[] Pdf, List<OoxPdfDiagnostic> Diagnostics) Convert(Dictionary<string, byte[]> parts)
    {
        using var input = TestFixtures.CreateZipPackage(parts);
        using var output = new MemoryStream();
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx,
            FontResolver = new PortableRenderingTests.PinnedResolver(), DiagnosticSink = diagnostics.Add });
        return (output.ToArray(), diagnostics);
    }

    private static Dictionary<string, byte[]> Parts(string? kind, string hidden, bool group)
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "pptx-images.pptx"));
        using var archive = ZipFile.OpenRead(path);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => { using var stream = entry.Open(); using var bytes = new MemoryStream(); stream.CopyTo(bytes); return bytes.ToArray(); });
        var original = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        var source = XDocument.Parse($$"""
            <p:sld xmlns:p="{{P}}" xmlns:a="{{A}}" showMasterSp="0"><p:cSld><p:spTree>
              <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
              <p:sp><p:nvSpPr><p:cNvPr id="2" name="Visible control"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></p:spPr>
                <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Visible control</a:t></a:r></a:p></p:txBody></p:sp>
            </p:spTree></p:cSld></p:sld>
            """);
        if (kind is not null)
        {
            XElement node;
            if (kind is "video" or "audio")
            {
                node = new XElement(original.Descendants(P + "pic").First());
                var properties = node.Element(P + "nvPicPr")!;
                properties.Element(P + "cNvPr")!.SetAttributeValue("id", 3);
                properties.Element(P + "cNvPr")!.SetAttributeValue("hidden", group ? "0" : hidden);
                if (properties.Element(P + "nvPr") is null) properties.Add(new XElement(P + "nvPr"));
                properties.Element(P + "nvPr")!.Add(new XElement(A + (kind + "File"), new XAttribute(R + "link", "probeMedia")));
                var relationships = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["ppt/slides/_rels/slide1.xml.rels"]));
                relationships.Root!.Add(new XElement(PackageR + "Relationship", new XAttribute("Id", "probeMedia"),
                    new XAttribute("Type", R.NamespaceName + "/" + kind), new XAttribute("Target", "file:///unused-public-media"), new XAttribute("TargetMode", "External")));
                parts["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8(relationships.ToString());
            }
            else
            {
                node = new XElement(P + "graphicFrame", new XElement(P + "nvGraphicFramePr",
                    new XElement(P + "cNvPr", new XAttribute("id", 3), new XAttribute("name", "Public frame"), new XAttribute("hidden", group ? "0" : hidden)),
                    new XElement(P + "cNvGraphicFramePr"), new XElement(P + "nvPr")),
                    new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", "urn:public:unsupported-frame"), kind == "ole" ? new XElement(P + "oleObj") : null)));
            }
            if (group)
                node = new XElement(P + "grpSp", new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr",
                    new XAttribute("id", 4), new XAttribute("name", "Public group"), new XAttribute("hidden", hidden)), new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")),
                    new XElement(P + "grpSpPr", new XElement(A + "xfrm",
                        new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                        new XElement(A + "ext", new XAttribute("cx", 9144000), new XAttribute("cy", 6858000)),
                        new XElement(A + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)),
                        new XElement(A + "chExt", new XAttribute("cx", 9144000), new XAttribute("cy", 6858000)))), node);
            source.Root!.Element(P + "cSld")!.Element(P + "spTree")!.Add(node);
        }
        parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(source.ToString());
        return parts;
    }
}
