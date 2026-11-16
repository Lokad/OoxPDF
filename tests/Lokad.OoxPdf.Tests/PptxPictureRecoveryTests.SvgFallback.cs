using System.Buffers.Binary;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxPictureRecoveryTests
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void PptxPictureRecoverySvgFallbackMatchesItsUsableSource()
    {
        foreach (bool cropped in new[] { false, true })
        {
            byte[] control = ConvertFallback(FallbackParts("svg-only", cropped: cropped), new RecoveryFontResolver(), out var controlDiagnostics);
            TestAssert.Equal(0, controlDiagnostics.Count);
            foreach (string variant in new[] { "relationship", "missing", "malformed", "payload" })
            {
                byte[] actual = ConvertFallback(FallbackParts(variant, cropped: cropped), new RecoveryFontResolver(), out var diagnostics);
                TestAssert.True(control.AsSpan().SequenceEqual(actual), "An unusable preview preserves the independently usable SVG's complete PDF.");
                var diagnostic = diagnostics.Single();
                TestAssert.Equal(OoxPdfSeverity.Warning, diagnostic.Severity);
                TestAssert.Equal("Embedded SVG", diagnostic.Fallback);
                TestAssert.Equal(1, diagnostic.SlideIndex);
                TestAssert.Equal(variant == "relationship" ? "IMAGE_MISSING_RELATIONSHIP" : variant == "missing" ? "IMAGE_MISSING_PART" : "IMAGE_UNSUPPORTED_FORMAT", diagnostic.Id);
            }
        }
        ConvertFallback(FallbackParts("both-missing"), new RecoveryFontResolver(), out var missing);
        TestAssert.Equal("Ignored", missing.Single().Fallback);
        TestAssert.Equal(OoxPdfSeverity.Error, missing.Single().Severity);
    }

    public static void PptxPictureRecoverySvgFontsFollowTheSelectedRepresentation()
    {
        var unused = new RecoveryFontResolver();
        byte[] preview = ConvertFallback(FallbackParts("valid", text: true), unused, out var diagnostics);
        byte[] raster = ConvertFallback(FallbackParts("raster-only", text: true), new RecoveryFontResolver(), out _);
        TestAssert.True(preview.AsSpan().SequenceEqual(raster), "Two usable sources preserve raster-preview precedence.");
        TestAssert.Equal(0, diagnostics.Count);
        TestAssert.True(!unused.Families.Contains("RecoverySvg"), "An unused SVG cannot trigger font acquisition.");
        var selected = new RecoveryFontResolver();
        byte[] recovered = ConvertFallback(FallbackParts("payload", text: true), selected, out _);
        byte[] svg = ConvertFallback(FallbackParts("svg-only", text: true), new RecoveryFontResolver(), out var svgDiagnostics);
        TestAssert.Equal(0, svgDiagnostics.Count);
        TestAssert.True(selected.Families.Contains("RecoverySvg"), "A recovered SVG prepares its selected text fonts before painting.");
        TestAssert.True(svg.AsSpan().SequenceEqual(recovered), "Recovery preserves glyph outlines as well as vector shapes.");
    }

    public static void PptxPictureRecoverySvgPreparationRespectsHiddenGroups()
    {
        foreach (bool group in new[] { false, true })
        {
            var parts = FallbackParts("payload", text: true);
            var slide = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
            var picture = slide.Descendants(P + "pic").Single();
            if (!group) picture.Descendants(P + "cNvPr").Single().SetAttributeValue("hidden", "1");
            else
            {
                var wrapper = new XElement(P + "grpSp", new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr", new XAttribute("id", 3), new XAttribute("name", "Group"), new XAttribute("hidden", "true")),
                    new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")), new XElement(P + "grpSpPr"), new XElement(picture));
                picture.ReplaceWith(wrapper);
            }
            parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide.ToString());
            var fonts = new RecoveryFontResolver();
            ConvertFallback(parts, fonts, out var diagnostics);
            TestAssert.Equal(0, diagnostics.Count);
            TestAssert.True(!fonts.Families.Contains("RecoverySvg"), "Hidden nodes cannot prepare fallback fonts.");
        }
    }

    public static void PptxPictureRecoverySvgDiagnosticsPreserveCancellationAndObservers()
    {
        foreach (bool cancel in new[] { false, true })
        {
            string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-svg-recovery-" + Guid.NewGuid().ToString("N"))).FullName;
            string input = Path.Combine(directory, "input.pptx"), output = Path.Combine(directory, "output.pdf");
            using var package = TestFixtures.CreateZipPackage(FallbackParts("payload"));
            File.WriteAllBytes(input, package.ToArray()); File.WriteAllText(output, "existing destination");
            var pool = new TestMemoryBudget(); var expected = new InvalidOperationException("observer failed");
            using var cancellation = new CancellationTokenSource();
            bool observed = false;
            try
            {
                void Observe(OoxPdfDiagnostic diagnostic)
                {
                    if (diagnostic.Fallback != "Embedded SVG") return;
                    observed = true;
                    if (cancel) cancellation.Cancel(); else throw expected;
                }
                void Convert() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = new RecoveryFontResolver(), MemoryBudget = pool, DiagnosticSink = Observe }, cancellation.Token);
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Convert).CancellationToken);
                else TestAssert.True(ReferenceEquals(expected, TestAssert.Throws<InvalidOperationException>(Convert)), "The observer retains its exception identity.");
                TestAssert.True(observed, "The recovery diagnostic must be delivered.");
                TestAssert.Equal("existing destination", File.ReadAllText(output));
                TestAssert.Equal(0, Directory.GetFiles(directory, "output.pdf.tmp-*").Length);
                TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            }
            finally { Directory.Delete(directory, recursive: true); }
        }
    }

    public static void PptxPictureRecoverySvgDoesNotRecoverHostAdmissionRefusal()
    {
        var parts = FallbackParts("valid");
        parts["ppt/media/preview.png"] = TestFixtures.CreateRgbPng(512, 512, new byte[512 * 512 * 3]);
        parts["ppt/slides/slide2.xml"] = parts["ppt/slides/slide1.xml"];
        parts["ppt/slides/_rels/slide2.xml.rels"] = parts["ppt/slides/_rels/slide1.xml.rels"];
        parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(PptxTests.InheritedShapePart("000000", 0, 0));
        var presentation = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/presentation.xml"]));
        presentation.Descendants(P + "sldIdLst").Single().Add(new XElement(P + "sldId", new XAttribute("id", 257), new XAttribute(R + "id", "second")));
        parts["ppt/presentation.xml"] = TestFixtures.Utf8(presentation.ToString());
        var relationships = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/_rels/presentation.xml.rels"]));
        relationships.Root!.Add(new XElement(relationships.Root.Name.Namespace + "Relationship", new XAttribute("Id", "second"), new XAttribute("Type", R.NamespaceName + "/slide"), new XAttribute("Target", "slides/slide2.xml")));
        parts["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(relationships.ToString());
        var types = XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/ppt/slides/slide2.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")));
        parts["[Content_Types].xml"] = TestFixtures.Utf8(types.ToString());
        var pool = new TestMemoryBudget(); var storage = new RecoveryRefusalStorage(pool); var diagnostics = new List<OoxPdfDiagnostic>();
        using var input = TestFixtures.CreateZipPackage(parts); using var output = new MemoryStream();
        var error = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = new RecoveryFontResolver(), MemoryBudget = pool, TemporaryStorage = storage,
                ConversionLimits = new() { MaxResidentPageContentBytesPerConversion = 0 }, DiagnosticSink = diagnostics.Add }));
        TestAssert.True(storage.Armed, "The first page arms host refusal before the picture's preparation.");
        TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, error.Kind);
        TestAssert.True(!diagnostics.Any(d => d.Fallback == "Embedded SVG"), "Host refusal cannot select image recovery.");
        TestAssert.Equal(0L, output.Length); TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    private static Dictionary<string, byte[]> FallbackParts(string variant, bool text = false, bool cropped = false)
    {
        string embed = variant == "svg-only" ? "svg" : "raster";
        string extension = variant is "svg-only" or "raster-only" ? "" : "<a:extLst><a:ext uri=\"svg\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"svg\"/></a:ext></a:extLst>";
        string crop = cropped ? "<a:srcRect l=\"10000\" r=\"10000\"/>" : "";
        var parts = new Dictionary<string, byte[]> {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace("</Types>", "<Default Extension=\"png\" ContentType=\"image/png\"/><Default Extension=\"svg\" ContentType=\"image/svg+xml\"/></Types>", StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8($"<p:sld xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id=\"2\" name=\"Picture\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{embed}\">{extension}</a:blip>{crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"1270000\" cy=\"1270000\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:ln><a:noFill/></a:ln></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8($"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + (variant == "relationship" ? "" : $"<Relationship Id=\"raster\" Type=\"{R}/image\" Target=\"../media/preview.png\"/>") + $"<Relationship Id=\"svg\" Type=\"{R}/image\" Target=\"../media/vector.svg\"/></Relationships>"),
            ["ppt/media/preview.png"] = TestFixtures.CreateTinyStoredRgbPng(),
            ["ppt/media/vector.svg"] = TestFixtures.Utf8(text ? "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100\" height=\"100\" viewBox=\"0 0 100 100\"><text x=\"2\" y=\"40\" font-family=\"RecoverySvg\" font-size=\"18\">LABEL</text></svg>" : "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100\" height=\"100\" viewBox=\"0 0 100 100\"><rect width=\"100\" height=\"100\" fill=\"#ff0000\"/></svg>")
        };
        if (variant is "missing" or "both-missing") parts.Remove("ppt/media/preview.png");
        if (variant == "both-missing") parts.Remove("ppt/media/vector.svg");
        if (variant == "malformed") parts["ppt/media/preview.png"][0] = 0;
        if (variant == "payload")
        {
            byte[] png = parts["ppt/media/preview.png"];
            for (int offset = 8; offset + 12 <= png.Length; offset += 12 + BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(offset, 4)))
                if (png.AsSpan(offset + 4, 4).SequenceEqual("IDAT"u8)) { png[offset + 8] = 0; break; }
        }
        return parts;
    }

    private static byte[] ConvertFallback(Dictionary<string, byte[]> parts, IFontResolver fonts, out List<OoxPdfDiagnostic> diagnostics)
    {
        var observed = new List<OoxPdfDiagnostic>();
        using var input = TestFixtures.CreateZipPackage(parts); using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = fonts, DiagnosticSink = observed.Add });
        diagnostics = observed; return output.ToArray();
    }

    private sealed class RecoveryFontResolver : IFontResolver
    {
        private readonly PortableRenderingTests.PinnedResolver inner = new();
        public readonly List<string> Families = [];
        public FontFaceResolution Resolve(FontRequest request) { Families.Add(request.FamilyName); return inner.Resolve(request); }
    }

    private sealed class RecoveryRefusalStorage(TestMemoryBudget pool) : IOoxPdfTemporaryStorage
    {
        public bool Armed;
        public ValueTask<Stream> CreateAsync(IMemoryBudget? memoryBudget, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            pool.Capacity = pool.Reserved; Armed = true;
            return ValueTask.FromResult<Stream>(new MemoryStream());
        }
    }
}
