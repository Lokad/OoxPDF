using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxImagesTests
{
    private static string PictureGeometryCase(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "PictureGeometry", name + ".pptx"));

    public static void PptxPictureGeometrySceneRetainsPresetAndAdjustmentOwners()
    {
        string path = PictureGeometryCase("picture-geometry-controls");
        TestAssert.Equal("E6130C6E1BF4382CDDDB1201866154EF7ACE0E053713288270BE927E1FE6DFD6",
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        using var stream = File.OpenRead(path);
        var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        var scene = new PptxSceneBuilder().Build(document, package, default);
        var expected = new[] { "rect", "ellipse", "roundRect", "roundRect", "triangle", "ellipse", "ellipse" };
        for (int i = 0; i < expected.Length; i++)
        {
            var picture = scene.Slides[i].SlideNodes.Single().Picture!;
            TestAssert.Equal(expected[i], picture.Preset);
            TestAssert.True(!picture.HasCustomGeometry, "Preset pictures retain their own geometry kind.");
            if (i == 3) TestAssert.Equal(40000d, picture.PresetAdjustments["adj"]);
        }
    }

    public static void PptxPictureGeometryPaintClipsTheAuthoredFrameAndKeepsSourceCrop()
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = RenderGeometryPages(PictureGeometryCase("picture-geometry-controls"), diagnostics);
        TestAssert.Equal(9, pages.Length);
        TestAssert.Equal(0, diagnostics.Count);
        // The ellipse passes through each frame midpoint, with four cubic arcs.
        // A rectangular viewport clip elsewhere on the page cannot satisfy this.
        foreach (int index in new[] { 1, 5, 6, 7 })
        {
            string content = NormalizeGeometryContent(pages[index].Content);
            TestAssert.Contains("360 288 m", content);
            TestAssert.Contains("311.647 396 252 396 c", content);
            TestAssert.Contains("144 347.647 144 288 c", content);
            TestAssert.Contains("192.353 180 252 180 c", content);
            TestAssert.Contains("360 228.353 360 288 c h W n", content);
            TestAssert.True(content.IndexOf("h W n", StringComparison.Ordinal) < content.IndexOf("/Im1 Do", StringComparison.Ordinal),
                "The authored mask is established before image paint.");
        }
        TestAssert.True(!NormalizeGeometryContent(pages[0].Content).Contains(" c ", StringComparison.Ordinal),
            "An authored rectangle retains its established viewport path.");
        TestAssert.Contains("252 396 m", NormalizeGeometryContent(pages[4].Content));
        TestAssert.Contains("360 180 l", NormalizeGeometryContent(pages[4].Content));
        TestAssert.Contains("144 180 l", NormalizeGeometryContent(pages[4].Content));
        TestAssert.True(pages[1].Images[0].Image.Width != pages[5].Images[0].Image.Width,
            "Source crop changes retained image pixels while leaving the frame mask intact.");
        TestAssert.True(NormalizeGeometryContent(pages[2].Content) != NormalizeGeometryContent(pages[3].Content),
            "Rounded-rectangle adjustment guides change the actual clip path.");
    }

    public static void PptxPictureGeometryNonSquareSvgGroupsAndFollowingPaintRetainFrames()
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = RenderGeometryPages(PictureGeometryCase("picture-geometry-extended"), diagnostics);
        TestAssert.Equal(4, pages.Length);
        TestAssert.True(diagnostics.Count == 0, string.Join("; ", diagnostics.Select(d => d.Id + ": " + d.Message)));
        foreach (int index in new[] { 0, 2, 3 })
        {
            string content = NormalizeGeometryContent(pages[index].Content);
            TestAssert.Contains("432 324 m", content);
            TestAssert.Contains("144 363.765 144 324 c", content);
            TestAssert.Contains("432 284.235 432 324 c h W n", content);
        }
        TestAssert.Equal(0, pages[2].Images.Count); // native SVG paint, not the raster branch
        foreach (var page in pages)
        {
            var states = new Stack<bool>(); bool masked = false, sawMask = false, sawFollowing = false;
            foreach (string line in page.Content.Split('\n'))
            {
                string item = line.Trim();
                if (item == "q") states.Push(masked);
                else if (item == "Q") { TestAssert.True(states.Count != 0, "Picture state restores without underflow."); masked = states.Pop(); }
                else if (item == "W n") { masked = true; sawMask = true; }
                else if (item.StartsWith("0.667 0 0.667", StringComparison.Ordinal))
                {
                    TestAssert.True(sawMask && !masked, "Following corner paint is outside the picture mask's graphics state.");
                    sawFollowing = true;
                }
            }
            TestAssert.True(sawFollowing && states.Count == 0 && !masked, "The following object paints after a complete mask-state restore.");
        }
    }

    public static void PptxPictureGeometryUnknownCustomAndBlurFallbacksRemainExplicit()
    {
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (string variant in new[] { "missing", "unknown", "custom", "outline", "shadow", "blur", "polygon-blur" })
        {
            var entries = new Dictionary<string, byte[]>();
            using (var archive = ZipFile.OpenRead(PictureGeometryCase("picture-geometry-controls")))
                foreach (var entry in archive.Entries)
                { using var bytes = new MemoryStream(); using var input = entry.Open(); input.CopyTo(bytes); entries[entry.FullName] = bytes.ToArray(); }
            var slide = XDocument.Parse(System.Text.Encoding.UTF8.GetString(entries["ppt/slides/slide1.xml"]));
            var props = slide.Descendants(p + "pic").Single().Element(p + "spPr")!;
            var preset = props.Element(a + "prstGeom")!;
            if (variant == "missing") preset.Remove();
            else if (variant == "custom") { preset.Remove(); props.Add(new XElement(a + "custGeom")); }
            else if (variant == "unknown") preset.SetAttributeValue("prst", "unimplementedPictureMask");
            else if (variant is "shadow" or "blur" or "polygon-blur")
            {
                preset.SetAttributeValue("prst", variant == "polygon-blur" ? "triangle" : "ellipse");
                props.Add(new XElement(a + "effectLst", new XElement(a + "outerShdw", new XAttribute("blurRad", variant is "blur" or "polygon-blur" ? 12700 : 0),
                    new XAttribute("dist", 12700), new XAttribute("dir", 5400000), new XElement(a + "srgbClr", new XAttribute("val", "000000")))));
            }
            else { preset.SetAttributeValue("prst", "ellipse"); props.Element(a + "ln")!.ReplaceWith(new XElement(a + "ln", new XAttribute("w", 12700), new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "000000"))))); }
            entries["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide.ToString());
            string path = TestFixtures.WriteTempPackage(".pptx", entries);
            try
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = RenderGeometryPages(path, diagnostics, 1);
                var relevant = diagnostics.Where(d => d.SlideIndex == 1).ToArray();
                string expected = variant == "polygon-blur" ? "PPTX_PICTURE_GEOMETRY_EFFECT_FALLBACK" : "PPTX_PICTURE_GEOMETRY_UNSUPPORTED";
                TestAssert.Equal(variant is "missing" or "outline" or "shadow" or "blur" ? 0 : 1, relevant.Length);
                if (relevant.Length != 0) TestAssert.Equal(expected, relevant[0].Id);
                bool hasEllipse = NormalizeGeometryContent(pages[0].Content).Contains("360 288 m", StringComparison.Ordinal);
                TestAssert.Equal(variant is "outline" or "shadow" or "blur", hasEllipse);
                using var canceled = new CancellationTokenSource(); canceled.Cancel();
                using var stream = File.OpenRead(path); var package = OoxPackage.Open(stream, default); var document = new PptxReader().Read(package, default);
                TestAssert.Throws<OperationCanceledException>(() => new PptxRenderer(null).RenderPages(document, package, null, canceled.Token).ToArray());
            }
            finally { File.Delete(path); }
        }
    }

    public static void PptxPictureGeometryVectorEffectsRetainIndependentOfficeExtents()
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = RenderGeometryPages(PictureGeometryCase("picture-geometry-effects"), diagnostics);
        TestAssert.Equal(6, pages.Length);
        TestAssert.Equal(0, diagnostics.Count);
        // Independent Word outline centerline bounds: x=142..434, y=250..398,
        // 4pt stroke around the authored x=144..432, y=252..396 image frame.
        string outline = NormalizeGeometryContent(pages[0].Content);
        TestAssert.Contains("434 324 m", outline);
        TestAssert.Contains("368.634 398 288 398 c", outline);
        TestAssert.Contains("142 364.869 142 324 c", outline);
        TestAssert.Contains("207.366 250 288 250 c", outline);
        TestAssert.Contains("434 283.131 434 324 c h S", outline);
        TestAssert.True(outline.IndexOf("434 324 m", StringComparison.Ordinal) > outline.IndexOf("/Im1 Do", StringComparison.Ordinal),
            "The complete external outline paints after the image mask restores.");
        string rounded = NormalizeGeometryContent(pages[1].Content);
        TestAssert.True(Regex.Matches(rounded, @" c ").Count == 8 && rounded.Contains(" h S", StringComparison.Ordinal),
            "Adjusted rounded picture mask and outline both retain curved corners.");
        string shadow = NormalizeGeometryContent(pages[3].Content);
        TestAssert.Contains("437.091 318.909 m", shadow);
        TestAssert.Contains("149.091 279.144 213.562 246.909 293.091 246.909 c", shadow);
        TestAssert.True(shadow.IndexOf("437.091 318.909 m", StringComparison.Ordinal) < shadow.IndexOf("432 324 m", StringComparison.Ordinal),
            "The offset ellipse shadow paints behind the unchanged authored picture mask.");
        TestAssert.True(pages.All(page => NormalizeGeometryContent(page.Content).Contains("144 376 20 20 re f*", StringComparison.Ordinal)),
            "Following corner markers remain painted after outlines, transforms and shadows.");
    }

    public static void PptxPictureGeometryBlurMasksKeepCornersTransparentAndRespectBudgets()
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var effects = RenderGeometryPages(PictureGeometryCase("picture-geometry-effects"), diagnostics);
        TestAssert.Equal(0, diagnostics.Count);
        var extra = RenderGeometryPages(PictureGeometryCase("picture-geometry-blur"), diagnostics);
        TestAssert.Equal(1, diagnostics.Count);
        TestAssert.Equal("PPTX_PICTURE_GEOMETRY_EFFECT_FALLBACK", diagnostics[0].Id);
        TestAssert.Equal(4, diagnostics[0].SlideIndex!.Value);
        foreach (var page in new[] { effects[5], extra[0], extra[1], extra[2] })
        {
            var shadow = page.Images[0].Image;
            TestAssert.True(shadow.Alpha is not null && shadow.Width <= 2048 && shadow.Height <= 2048,
                "The shape shadow uses the existing bounded effect raster and soft mask.");
            using var bytes = new MemoryStream(shadow.Alpha!);
            using var inflater = new ZLibStream(bytes, CompressionMode.Decompress);
            using var pixels = new MemoryStream(); inflater.CopyTo(pixels);
            byte[] alpha = pixels.ToArray();
            TestAssert.Equal(shadow.Width * shadow.Height, alpha.Length);
            TestAssert.Equal((byte)0, alpha[shadow.Height / 20 * shadow.Width + shadow.Width / 20]);
            TestAssert.Equal((byte)255, alpha[shadow.Height / 2 * shadow.Width + shadow.Width / 2]);
            TestAssert.Contains("144 376 20 20 re f*", NormalizeGeometryContent(page.Content));
        }
        foreach (var limits in new[] {
            new OoxConversionLimits { MaxRetainedImageBytesPerConversion = 0 },
            new OoxConversionLimits { MaxLiveImageBytesPerConversion = 0 } })
        {
            using var input = File.OpenRead(PictureGeometryCase("picture-geometry-blur"));
            using var output = new MemoryStream();
            TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
                new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, ConversionLimits = limits }));
            TestAssert.Equal(0L, output.Length);
        }
    }

    public static void PptxPictureGeometryCssKeywordsAndInvalidHexNeverThrow()
    {
        foreach (var pair in new[] { ("yellow", new RgbColor(255, 255, 0)), ("orange", new RgbColor(255, 165, 0)),
            ("purple", new RgbColor(128, 0, 128)), ("silver", new RgbColor(192, 192, 192)), ("maroon", new RgbColor(128, 0, 0)) })
        {
            TestAssert.True(RgbColor.TryParseCssColor(pair.Item1, out var color), "A six-letter CSS keyword is not a malformed hex token.");
            TestAssert.Equal(pair.Item2, color);
        }
        foreach (string invalid in new[] { "yellow", "GG0000", "00xx00", "0000zz", "-10000" })
            TestAssert.True(!RgbColor.TryParse(invalid, out _), "Invalid hex reports false instead of aborting the enclosing picture.");
        foreach (string invalid in new[] { "#GG0000", "#xyz", "nocolor" })
            TestAssert.True(!RgbColor.TryParseCssColor(invalid, out _), "Malformed CSS paint retains the explicit unsupported-paint path.");
        TestAssert.True(RgbColor.TryParseCssColor("#fF0", out var shortColor) && shortColor == new RgbColor(255, 255, 0), "Safe hex parsing preserves supported shorthand.");
    }

    private static PdfPage[] RenderGeometryPages(string path, List<OoxPdfDiagnostic> diagnostics, int count = int.MaxValue)
    {
        using var stream = File.OpenRead(path);
        var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(null).RenderPages(document, package, diagnostics.Add, default).Take(count).ToArray();
    }

    private static string NormalizeGeometryContent(string content) => Regex.Replace(content, @"\s+", " ").Trim();
}
