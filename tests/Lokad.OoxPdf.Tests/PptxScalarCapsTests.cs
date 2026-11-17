using System.Buffers.Binary;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxScalarCapsTests
{
    public static void PptxScalarCapsClassifiesCompleteLettersAndPreservesFragmentScales()
    {
        const string source = "a\U00010428Z\U00010400\U000104D8\U000104B0!";
        var actual = Inspect(source, "small");
        int[] expected = [65, 0x10400, 90, 0x10400, 0x104B0, 0x104B0, 33];
        double[] sizes = [19.2, 19.2, 24, 24, 19.2, 24, 24];
        TestAssert.True(actual.Select(item => item.CodePoint).SequenceEqual(expected), "Deseret and Osage lowercase letters uppercase as whole scalars.");
        TestAssert.True(actual.Select(item => item.Size).Zip(sizes).All(pair => Math.Abs(pair.First - pair.Second) < .000001d),
            "Only originally lowercase letters receive the existing small-caps scale.");
        TestAssert.Equal(expected.Length, actual.Length);
    }

    public static void PptxScalarCapsModesRetainUncasedScalarsAndDirectOverrides()
    {
        const string source = "a\U00010428\U000104D8\U0001F600";
        int[] original = [97, 0x10428, 0x104D8, 0x1F600];
        int[] uppercase = [65, 0x10400, 0x104B0, 0x1F600];
        foreach (string mode in new[] { "none", "all", "small" })
        {
            var actual = Inspect(source, mode);
            TestAssert.True(actual.Select(item => item.CodePoint).SequenceEqual(mode == "none" ? original : uppercase), "Case mode retains every source scalar, including an uncased symbol.");
            for (int index = 0; index < actual.Length; index++)
                TestAssert.True(Math.Abs(actual[index].Size - (mode == "small" && index < 3 ? 19.2d : 24d)) < .000001d,
                    "Scalar width and casing category are separate from nominal font size.");
        }
        TestAssert.True(Inspect(source, null, "small").SequenceEqual(Inspect(source, "small")), "An inherited cap mode uses the same fragment rule.");
        TestAssert.True(Inspect(source, "none", "small").SequenceEqual(Inspect(source, "none")), "Explicit none stops inherited casing and scaling.");
    }

    public static void PptxScalarCapsEmissionKeepsEveryTransformedSourceLetter()
    {
        using var input = Input("a\U00010428\U000104D8\U0001F600", "small");
        using var package = OoxPackage.Open(input, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = new PptxRenderer(new ScalarResolver()).RenderPagesAsync(new PptxReader().Read(package, default), package,
            diagnostics.Add, default).ToBlockingEnumerable().Single();
        TestAssert.Equal("A\U00010400\U000104B0\U0001F600", DocxPageFootnoteNumberingTests.DecodeText(page));
        TestAssert.Contains("19.2 Tf", page.Content);
        TestAssert.Contains("24 Tf", page.Content);
        TestAssert.Equal(0, diagnostics.Count);
        var encoded = page.Fonts.SelectMany(font => font.Font.UnicodeByCid.Values).ToHashSet();
        TestAssert.True(encoded.Contains(0x10400) && encoded.Contains(0x104B0) && !encoded.Contains(0x10428) && !encoded.Contains(0x104D8),
            "Preparation and emission own the transformed characters, not stale lowercase outlines.");
    }

    private static (int CodePoint, double Size)[] Inspect(string text, string? caps, string? inheritedCaps = null)
    {
        using var input = Input(text, caps, inheritedCaps);
        using var package = OoxPackage.Open(input, default);
        var layout = PptxRenderer.InspectTextLayout(new PptxReader().Read(package, default), package, 0,
            inspectionFonts: new PresentationFontResolver(new ScalarResolver()));
        return layout.Frames.SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans).SelectMany(span => span.GlyphSpan.Glyphs.Select(glyph => (glyph.CodePoint, span.FontSize))).ToArray();
    }

    private static MemoryStream Input(string text, string? caps, string? inheritedCaps = null)
    {
        XNamespace a = OoxNamespaces.DrawingNamespace;
        var run = new XElement(a + "r", new XElement(a + "rPr", new XAttribute("sz", "2400"),
            caps is null ? null : new XAttribute("cap", caps), new XElement(a + "latin", new XAttribute("typeface", "Scalar control"))),
            new XElement(a + "t", text));
        var paragraph = new XElement(a + "p", inheritedCaps is null ? null : new XElement(a + "pPr",
            new XElement(a + "defRPr", new XAttribute("cap", inheritedCaps))), run);
        return TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp>
                  <p:spPr><a:xfrm><a:off x="254000" y="254000"/><a:ext cx="7620000" cy="1270000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>{{paragraph}}</p:txBody>
                </p:sp></p:spTree></p:cSld></p:sld>
                """
        });
    }

    private sealed class ScalarResolver : IFontResolver
    {
        private readonly MemoryFontProgramSource program = new("scalar-case-control", FontBytes());
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "Scalar control", new FontStyleKey(), program, false);
    }

    private static byte[] FontBytes()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "PortableFonts", "DejaVuSans.ttf"));
        byte[] bytes = File.ReadAllBytes(path);
        var font = OpenTypeFont.Load(bytes);
        var points = Enumerable.Range(32, 95).Select(point => (Point: point, Glyph: font.MapCodePoint(point))).ToList();
        // Distinct A/a and B/b outlines prove both case selection and scale;
        // source Unicode identities remain distinct even when outlines are shared.
        points.AddRange([(0x10400, font.MapCodePoint('A')), (0x10428, font.MapCodePoint('a')),
            (0x104B0, font.MapCodePoint('B')), (0x104D8, font.MapCodePoint('b')), (0x1F600, font.MapCodePoint('!'))]);
        TestAssert.True(points.All(item => item.Glyph != 0), "Every controlled scalar has explicit font coverage.");
        TestAssert.True(points.Single(item => item.Point == 0x10400).Glyph != points.Single(item => item.Point == 0x10428).Glyph,
            "Uppercase/lowercase controls must use different glyph outlines.");
        byte[] cmap = new byte[28 + 12 * points.Count];
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(2), 1);
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(4), 3);
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(6), 10);
        BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(8), 12);
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(12), 12);
        BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(16), (uint)cmap.Length - 12);
        BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(24), (uint)points.Count);
        for (int index = 0; index < points.Count; index++)
        {
            int offset = 28 + 12 * index;
            BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(offset), (uint)points[index].Point);
            BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(offset + 4), (uint)points[index].Point);
            BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(offset + 8), points[index].Glyph);
        }
        return TestFontBuilder.ReplaceTable(bytes, "cmap", cmap);
    }
}
