using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class FallbackFontMetricsTests
{
    private const string Prefix = "WWW MMM iii";

    public static void BuiltInFallbackAdvancesMatchPrimaryFontMetrics()
    {
        foreach (bool bold in new[] { false, true })
        foreach (bool italic in new[] { false, true })
        {
            var face = PdfFallbackFont.ForStyle(bold, italic);
            var metrics = ReadMetrics(face.FaceName);
            foreach (char letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
                TestAssert.Equal((double)metrics[letter.ToString()], face.MeasureAdvanceEm(letter));
            string[] digits = "zero one two three four five six seven eight nine".Split(' ');
            for (int i = 0; i < digits.Length; i++)
                TestAssert.Equal((double)metrics[digits[i]], face.MeasureAdvanceEm('0' + i));
            foreach (var (point, name) in new (int, string)[] {
                (' ', "space"), ('?', "question"), ('\'', "quotesingle"), ('`', "grave"),
                (0xa0, "space"), (0xb2, "twosuperior"), (0xb3, "threesuperior"), (0xb9, "onesuperior"),
                (0xe9, "eacute"), (0xff, "ydieresis"), (0x20ac, "Euro"), (0x2014, "emdash"),
                (0x2019, "quoteright"), (0x0153, "oe"), (0x0178, "Ydieresis"), (0x2122, "trademark") })
                TestAssert.Equal((double)metrics[name], face.MeasureAdvanceEm(point));
        }
    }

    public static void FallbackPositioningRetainsAllocationSubstitutionAndCancellation()
    {
        foreach (bool bold in new[] { false, true })
        {
            var face = PdfFallbackFont.ForStyle(bold, italic: false);
            var metrics = ReadMetrics(face.FaceName);
            var glyphs = face.PositionGlyphs("A\tĀ\u0301中B", 20d, 40d, 10d, 1d);
            TestAssert.Equal("65,63,63,63,66", string.Join(',', glyphs.Select(g => g.Code)));
            TestAssert.Equal(0d, face.MeasureAdvanceEm('\t'));
            TestAssert.True(Math.Abs(glyphs[1].X - glyphs[0].X - metrics["A"] * .01d - 2d) < 1e-9,
                "Nonpainting controls retain tracking; unsupported combining and wide characters advance as their emitted question marks.");
            for (int index = 1; index < 4; index++)
                TestAssert.True(Math.Abs(glyphs[index + 1].X - glyphs[index].X - metrics["question"] * .01d - 1d) < 1e-9,
                    "Every emitted substitution owns its declared advance.");
            var allocated = face.PositionGlyphs("A\tĀ\u0301中B", 20d, 40d, 10d, 1d, allocatedWidth: 180d);
            double units = metrics["A"] + 3d * metrics["question"] + metrics["B"];
            TestAssert.True(Math.Abs(allocated[^1].X + metrics["B"] * 175d / units - 200d) < 1e-9,
                "PPTX allocations preserve the final edge and tracking while distributing proportional advances.");
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Equal(canceled.Token, TestAssert.Throws<OperationCanceledException>(() =>
            PdfFallbackFont.Helvetica.PositionGlyphs(Prefix, 0d, 0d, 11d, 0d, token: canceled.Token)).CancellationToken);
    }

    public static void FallbackGlyphEmissionObservesCancellationDuringConstruction()
    {
        var budget = new TestMemoryBudget();
        using var cancellation = new CancellationTokenSource();
        using (var scope = budget.CreateScope())
        using (var grant = OoxMemory.Reserve(scope, 0)!)
        {
            var graphics = new PdfGraphicsBuilder(grant, 0);
            var glyphs = PdfFallbackFont.Helvetica.PositionGlyphs(new string('W', 512), 0d, 20d, 12d, 0d);
            budget.RefuseChange = (before, after) =>
            {
                if (after > before && graphics.ContentLength > 0)
                {
                    cancellation.Cancel();
                }
                return false;
            };
            var error = TestAssert.Throws<OperationCanceledException>(() =>
                graphics.DrawFallbackText("F0", 12d, 0, 0, 0, glyphs, cancellation.Token));
            TestAssert.Equal(cancellation.Token, error.CancellationToken);
            TestAssert.True(graphics.ContentLength > 0, "Cancellation must be observed during fallback glyph emission.");
        }
        TestAssert.Equal(0L, budget.Reserved);
    }

    public static void DocxFallbackGlyphPaintAndWrapUseDeclaredFaceWidths()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            foreach (bool bold in new[] { false, true })
            foreach (bool italic in new[] { false, true })
            {
                var source = ReadDocx(mode, Prefix + " WidthProbe ", bold, italic);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(new MissingTestFontResolver(), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(2, pages.Length);
                TestAssert.True(pages.All(page => page.Fonts.Count == 0), "Exercise standard-font paint, not an embedded substitute.");
                var glyphs = ReadGlyphs(pages[0].Content).Take(Prefix.Length).ToArray();
                string face = bold ? (italic ? "Helvetica-BoldOblique" : "Helvetica-Bold") : italic ? "Helvetica-Oblique" : "Helvetica";
                CheckPrefix(glyphs, ReadMetrics(face), tracking: 0d);
                TestAssert.Equal(face, pages[0].FallbackFonts.Single(f => f.ResourceName == glyphs[0].Font).Font.FaceName);
                TestAssert.True(diagnostics.Any(d => d.Id == "FONT_NO_USABLE_FACE"), "Fallback stays diagnosed.");
            }
            var wide = ReadDocx(mode, new string('W', 60) + " ", bold: false, italic: false);
            var wrapped = new DocxRenderer(new MissingTestFontResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(wide, null, default).ToBlockingEnumerable().ToArray();
            var letters = ReadGlyphs(wrapped[0].Content).Where(g => g.Code == (byte)'W').Take(60).ToArray();
            TestAssert.Equal(60, letters.Length);
            TestAssert.True(letters.Select(g => g.Y).Distinct().Count() >= 2,
                "A word wider than the body frame wraps using real advances instead of fitting through overlapping glyphs.");
        }
    }

    public static void PptxFallbackGlyphPaintUsesDeclaredFaceWidths()
    {
        var shapes = new StringBuilder();
        foreach (var (bold, italic, index) in new[] { (false, false, 0), (true, false, 1), (false, true, 2), (true, true, 3) })
            shapes.Append(Shape(bold, italic, index));

        XElement contentTypes = XElement.Parse(PptxTests.BasicContentTypes());
        XElement presentation = XElement.Parse(PptxTests.BasicPresentation());
        XElement relationships = XElement.Parse(PptxTests.PresentationRelationship());
        XElement slideIds = presentation.Element(OoxNamespaces.PresentationNamespace + "sldIdLst")!;
        foreach (int index in new[] { 2, 3 })
        {
            string relationshipId = "rId" + index;
            contentTypes.Add(new XElement(OoxNamespaces.ContentTypesNamespace + "Override",
                new XAttribute("PartName", $"/ppt/slides/slide{index}.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")));
            slideIds.Add(new XElement(OoxNamespaces.PresentationNamespace + "sldId",
                new XAttribute("id", 255 + index), new XAttribute(OoxNamespaces.RelationshipsNamespace + "id", relationshipId)));
            relationships.Add(new XElement(OoxNamespaces.PackageRelationshipsNamespace + "Relationship",
                new XAttribute("Id", relationshipId), new XAttribute("Type", PptxReader.SlideRelationshipType),
                new XAttribute("Target", $"slides/slide{index}.xml")));
        }
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = contentTypes.ToString(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = relationships.ToString(),
            ["ppt/presentation.xml"] = presentation.ToString(),
            ["ppt/slides/slide1.xml"] = Slide(shapes.ToString()),
            ["ppt/slides/slide2.xml"] = Slide(""),
            ["ppt/slides/slide3.xml"] = Slide(Shape(bold: false, italic: false, index: 0))
        });
        using OoxPackage package = OoxPackage.Open(input, default);
        PptxDocument document = new PptxReader().Read(package, default);
        PdfPage[] pages = new PptxRenderer(new MissingTestFontResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(3, pages.Length);
        TestAssert.True(pages.All(page => page.Fonts.Count == 0), "Exercise standard-font paint.");
        for (int index = 0; index < pages.Length; index++)
            PdfContentValidator.ValidatePage(pages[index], index, default);

        TestAssert.Equal("FF1,FF2,FF3,FF4", string.Join(',', pages[0].FallbackFonts.Select(font => font.ResourceName)));
        TestAssert.Equal(0, pages[1].FallbackFonts.Count);
        TestAssert.Equal("FF1", pages[2].FallbackFonts.Single().ResourceName);
        CheckPrefix(ReadGlyphs(pages[2].Content).Take(Prefix.Length).ToArray(), ReadMetrics("Helvetica"), tracking: 1d);

        var glyphs = ReadGlyphs(pages[0].Content).ToArray();
        foreach (string font in new[] { "FF1", "FF2", "FF3", "FF4" })
        {
            var prefix = glyphs.Where(g => g.Font == font).Take(Prefix.Length).ToArray();
            string face = font == "FF1" ? "Helvetica" : font == "FF2" ? "Helvetica-Bold" : font == "FF3" ? "Helvetica-Oblique" : "Helvetica-BoldOblique";
            CheckPrefix(prefix, ReadMetrics(face), tracking: 1d);
        }

        static string Slide(string content) => $"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>{content}</p:spTree></p:cSld>
            </p:sld>
            """;

        static string Shape(bool bold, bool italic, int index) => FormattableString.Invariant($"""
            <p:sp><p:spPr><a:xfrm><a:off x="914400" y="{914400 + index * 1143000}"/><a:ext cx="7315200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
            <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none"/><a:lstStyle/><a:p><a:r>
            <a:rPr sz="2400" spc="100" b="{(bold ? 1 : 0)}" i="{(italic ? 1 : 0)}"><a:latin typeface="Absent"/></a:rPr><a:t>{Prefix} WidthProbe</a:t>
            </a:r></a:p></p:txBody></p:sp>
            """);
    }

    private static DocxDocument ReadDocx(OoxPdfDocxMarkupMode mode, string prefix, bool bold, bool italic) =>
        DocxInlineCompositionTests.ReadMutation("CompositionProbes/math-page-notes-arial-11-zero", mode, false, false, (xml, w, m) =>
        {
            xml.Descendants(m + "oMath").Remove();
            var text = xml.Descendants(w + "t").First(t => t.Value.Contains("Alpha before", StringComparison.Ordinal));
            text.Value = prefix + text.Value;
            var properties = text.Parent!.Element(w + "rPr");
            if (properties is null) { properties = new XElement(w + "rPr"); text.Parent.AddFirst(properties); }
            properties.Elements(w + "b").Remove(); properties.Elements(w + "i").Remove();
            properties.Add(new XElement(w + "b", new XAttribute(w + "val", bold ? "1" : "0")),
                new XElement(w + "i", new XAttribute(w + "val", italic ? "1" : "0")));
        });

    private static Dictionary<string, int> ReadMetrics(string face)
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CoreFonts", face + ".afm"));
        return Regex.Matches(File.ReadAllText(path), @"^C\s+-?\d+\s*;\s*WX\s+(\d+)\s*;\s*N\s+(\S+)\s*;", RegexOptions.Multiline)
            .ToDictionary(m => m.Groups[2].Value, m => int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture));
    }

    internal readonly record struct Glyph(string Font, double Size, double X, double Y, byte Code);

    internal static IEnumerable<Glyph> ReadGlyphs(string content)
    {
        foreach (Match region in Regex.Matches(content, @"BT(.*?)ET", RegexOptions.Singleline))
        {
            Match face = Regex.Match(region.Value, @"/(FF\d)\s+([\d.]+)\s+Tf");
            if (!face.Success) continue;
            double size = double.Parse(face.Groups[2].Value, CultureInfo.InvariantCulture);
            foreach (Match point in Regex.Matches(region.Value, @"1 0 0 1 ([-\d.]+) ([-\d.]+) Tm\s*<([0-9A-Fa-f]{2})> Tj"))
                yield return new(face.Groups[1].Value, size, double.Parse(point.Groups[1].Value, CultureInfo.InvariantCulture),
                    double.Parse(point.Groups[2].Value, CultureInfo.InvariantCulture), Convert.ToByte(point.Groups[3].Value, 16));
        }
    }

    private static void CheckPrefix(Glyph[] glyphs, Dictionary<string, int> metrics, double tracking)
    {
        TestAssert.Equal(Prefix, Encoding.ASCII.GetString(glyphs.Select(g => g.Code).ToArray()));
        for (int i = 0; i < glyphs.Length - 1; i++)
        {
            string name = Prefix[i] == ' ' ? "space" : Prefix[i].ToString();
            double expected = metrics[name] * glyphs[i].Size / 1000d + tracking;
            TestAssert.True(Math.Abs(glyphs[i + 1].X - glyphs[i].X - expected) <= .0011d,
                "Glyph origins must follow primary AFM advances plus authored tracking: " + Prefix[i]);
            TestAssert.Equal(glyphs[i].Y, glyphs[i + 1].Y);
        }
    }
}
