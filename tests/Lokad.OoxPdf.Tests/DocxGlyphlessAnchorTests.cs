using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxGlyphlessAnchorTests
{
    public static void DocxGlyphlessFloatingImagePaintsWithoutFontPrograms()
    {
        string input = CreateDocument();
        foreach (var (mode, geometry) in Views())
        {
            PdfPage page = Render(input, new EmptyProgramResolver(), mode, geometry).Single();
            TestAssert.Equal(0, page.Fonts.Count);
            double[] frame = PaintedFrame(page);
            double[] expected = [144d, 72d, 234d, page.Height - 149d];
            TestAssert.True(frame.Zip(expected).All(p => Math.Abs(p.First - p.Second) < .001d), "Fontless drawing placement keeps the source extent, centered column and5pt paragraph offset.");
        }
    }

    public static void DocxGlyphlessFloatingFramesMatchAnExplicitFontControl()
    {
        string input = CreateDocument();
        foreach (var (mode, geometry) in Views())
        {
            double[] missing = PaintedFrame(Render(input, new EmptyProgramResolver(), mode, geometry).Single());
            double[] available = PaintedFrame(Render(input, new PortableRenderingTests.PinnedResolver(), mode, geometry).Single());
            TestAssert.True(missing.Zip(available).All(p => Math.Abs(p.First - p.Second) < .001d), "An unavailable font does not change a glyphless floating picture's physical frame.");
        }
    }

    public static void DocxGlyphlessAnchorsFollowBlankSlotsAndPageBreaks()
    {
        var parts = ReadParts(CreateDocument());
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["word/document.xml"]));
        XElement anchor = xml.Descendants(w + "p").First();
        XElement Blank() => new(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
            new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"), new XAttribute(w + "line", "400"), new XAttribute(w + "lineRule", "exact"))));
        anchor.AddBeforeSelf(Blank(), new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page")))), Blank(), Blank());
        parts["word/document.xml"] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        PdfPage[] pages = Render(input, new EmptyProgramResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        TestAssert.Equal(2, pages.Length);
        TestAssert.Equal(0, pages[0].Images.Count);
        double[] frame = PaintedFrame(pages[1]);
        TestAssert.True(Math.Abs(pages[1].Height - frame[3] - frame[1] - 117d) < .001d, "Two authored20pt blank slots move the paragraph anchor40pt on its own page.");
        TestAssert.True(pages.All(p => p.Fonts.Count == 0), "Empty slots do not import a font program.");
    }

    public static void DocxGlyphlessPictureFilePublishesAndRefundsWithoutFonts()
    {
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        using var budget = new TestMemoryBudget();
        try
        {
            File.WriteAllBytes(output, [1, 2, 3]);
            OoxPdfConverter.Convert(CreateDocument(), output, new OoxPdfOptions { FontResolver = new EmptyProgramResolver(), MemoryBudget = budget });
            string pdf = File.ReadAllText(output, System.Text.Encoding.ASCII);
            TestAssert.Contains("/Subtype /Image", pdf);
            TestAssert.True(Regex.IsMatch(pdf, @"/[A-Za-z0-9]+ Do\b"), "The published PDF paints an XObject.");
            TestAssert.DoesNotContain("/FontFile", pdf);
            TestAssert.Equal(0L, budget.Reserved);
            TestAssert.True(budget.Peak > 0, "The caller funds the conversion's existing resource owners.");
            TestAssert.Equal(0, budget.CallerDisposals);
        }
        finally { File.Delete(output); }
    }

    public static void DocxGlyphlessPictureLateCancellationPreservesPublication()
    {
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        byte[] original = [1, 2, 3];
        using var cancellation = new CancellationTokenSource();
        using var budget = new TestMemoryBudget();
        bool completedRendering = false;
        try
        {
            File.WriteAllBytes(output, original);
            TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(CreateDocument(), output,
                new OoxPdfOptions { FontResolver = new EmptyProgramResolver(), MemoryBudget = budget, ReportResourceUsage = true,
                    DiagnosticSink = diagnostic => { if (diagnostic.Id == "CONVERSION_RESOURCE_SUMMARY") { completedRendering = true; cancellation.Cancel(); } } }, cancellation.Token));
            TestAssert.True(completedRendering, "Cancellation occurs after the fontless image page is rendered.");
            TestAssert.True(original.AsSpan().SequenceEqual(File.ReadAllBytes(output)), "Late cancellation preserves the published destination.");
            TestAssert.Equal(0L, budget.Reserved);
            TestAssert.Equal(0, budget.CallerDisposals);
        }
        finally { File.Delete(output); }
    }

    internal static string CreateDocument()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="114300" distB="228600" distL="342900" distR="457200"
                                     simplePos="0" relativeHeight="251658240" behindDoc="0"
                                     locked="1" layoutInCell="1" allowOverlap="0">
                            <wp:extent cx="1828800" cy="914400"/>
                            <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:posOffset>63500</wp:posOffset></wp:positionV>
                            <wp:wrapSquare wrapText="bothSides"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
    }

    private static (OoxPdfDocxMarkupMode, OoxPdfDocxMarkupGeometryMode)[] Views() =>
    [
        (OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout),
        (OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout),
        (OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
    ];

    private static PdfPage[] Render(string input, IFontResolver resolver, OoxPdfDocxMarkupMode mode, OoxPdfDocxMarkupGeometryMode geometry)
    {
        using var stream = File.OpenRead(input);
        using var package = OoxPackage.Open(stream, default);
        var document = new DocxReader().Read(package, null, default, markupMode: mode);
        return new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
    }

    private static double[] PaintedFrame(PdfPage page)
    {
        var image = page.Images.Single();
        var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" + Regex.Escape(image.ResourceName) + @" Do");
        TestAssert.True(draw.Success, "The picture resource must be painted.");
        return Enumerable.Range(1, 4).Select(i => double.Parse(draw.Groups[i].Value, CultureInfo.InvariantCulture)).ToArray();
    }

    private static Dictionary<string, byte[]> ReadParts(string input)
    {
        using var stream = File.OpenRead(input);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var source = entry.Open(); using var output = new MemoryStream(); source.CopyTo(output); return output.ToArray();
        });
    }

    private sealed class EmptyProgramResolver : IFontResolver
    {
        private static readonly MemoryFontProgramSource Source = new("glyphless-empty-font", ReadOnlyMemory<byte>.Empty);
        public FontFaceResolution Resolve(FontRequest request) =>
            new(request.FamilyName, "Helvetica", new FontStyleKey(request.Bold, request.Italic), Source, IsFallback: true);
    }
}
