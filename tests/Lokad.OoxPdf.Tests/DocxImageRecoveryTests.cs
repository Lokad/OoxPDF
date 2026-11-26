using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxImageRecoveryTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly OoxPdfDocxMarkupMode[] Modes = [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup];

    public static void DocxPictureMetadataPreservesRasterAndSvgPaint()
    {
        XNamespace svg = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
        foreach (var (file, storyParts, vector) in new[]
        {
            ("Composition/composed-static-cell-report.docx", new[] { "word/document.xml", "word/endnotes.xml", "word/footer1.xml" }, false),
            ("CompositionProbes/composed-static-svg-report.docx", new[] { "word/footer1.xml" }, true)
        })
        {
            string source = TestFixtures.PublicCasePath(file);
            Mutate(source, parts => {
                foreach (string part in storyParts)
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[part]));
                    var decoy = vector ? new XElement(svg + "svgBlip", new XAttribute(R + "embed", "ignoredPreview"))
                        : new XElement(A + "blip");
                    AddPicturePreview(xml.Descendants(Pic + "pic").Single(), decoy);
                    parts[part] = TestFixtures.Utf8(xml.ToString());
                }
            }, changed => {
                foreach (var mode in Modes)
                {
                    var baseline = ConvertPictureTest(source, mode);
                    var actual = ConvertPictureTest(changed, mode);
                    TestAssert.True(baseline.Pdf.AsSpan().SequenceEqual(actual.Pdf), $"Picture paint or layout changed in {file}, {mode}.");
                    TestAssert.True(baseline.Diagnostics.SequenceEqual(actual.Diagnostics), $"Picture diagnostics changed in {file}, {mode}.");
                }
            });
        }
    }

    public static void DocxPictureMetadataKeepsPreviewDrawingsAndWarningsOutOfOutput()
    {
        const string document = "word/document.xml";
        const string relationships = "word/_rels/document.xml.rels";
        string source = TestFixtures.PublicCasePath("ImageRecovery/picture-present.docx");
        Mutate(source, parts => {
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[document]));
            var inline = xml.Descendants(WP + "inline").Single();
            var graphic = inline.Element(A + "graphic")!;
            var anchor = new XElement(WP + "anchor",
                new XAttribute("simplePos", "0"), new XAttribute("relativeHeight", "1"),
                new XAttribute("behindDoc", "0"), new XAttribute("locked", "0"),
                new XAttribute("layoutInCell", "1"), new XAttribute("allowOverlap", "1"),
                new XElement(WP + "simplePos", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(WP + "positionH", new XAttribute("relativeFrom", "margin"), new XElement(WP + "posOffset", 0)),
                new XElement(WP + "positionV", new XAttribute("relativeFrom", "page"), new XElement(WP + "posOffset", 3810000)),
                inline.Element(WP + "extent"), new XElement(WP + "wrapNone"), inline.Element(WP + "docPr"), graphic);
            AddPicturePreview(graphic.Descendants(Pic + "pic").Single(), new XElement(inline), anchor,
                new XElement(A + "blip", new XAttribute(R + "link", "previewExternal")),
                new XElement(XName.Get("chart", "http://schemas.openxmlformats.org/drawingml/2006/chart")),
                new XElement(XName.Get("relIds", "http://schemas.openxmlformats.org/drawingml/2006/diagram")),
                new XElement(XName.Get("rect", "urn:schemas-microsoft-com:vml")), new XElement(W + "object"));
            parts[document] = TestFixtures.Utf8(xml.ToString());

            var rels = XDocument.Parse(Encoding.UTF8.GetString(parts[relationships]));
            rels.Root!.Add(new XElement(rels.Root.Name.Namespace + "Relationship", new XAttribute("Id", "previewExternal"),
                new XAttribute("Type", R.NamespaceName + "/image"), new XAttribute("TargetMode", "External"),
                new XAttribute("Target", "https://example.org/preview.png")));
            parts[relationships] = TestFixtures.Utf8(rels.ToString());
        }, changed => {
            foreach (var mode in Modes)
            {
                var baseline = ConvertPictureTest(source, mode);
                var actual = ConvertPictureTest(changed, mode);
                TestAssert.True(baseline.Pdf.AsSpan().IndexOf("/Subtype /Image"u8) >= 0, "The source picture must remain present.");
                TestAssert.True(baseline.Pdf.AsSpan().SequenceEqual(actual.Pdf), $"Copied drawing frames changed the PDF in {mode}.");
                TestAssert.Equal(0, baseline.Diagnostics.Count);
                TestAssert.Equal(0, actual.Diagnostics.Count);
            }
        });
    }

    private static void AddPicturePreview(XElement picture, params XElement[] payload)
    {
        picture.Element(Pic + "nvPicPr")!.Element(Pic + "cNvPr")!.Add(new XElement(A + "extLst", new XElement(A + "ext",
            new XAttribute("uri", "{5D154C70-1063-4818-8DFE-C915C125A6AD}"),
            new XElement(XName.Get("preview", "urn:ooxpdf:test"), payload))));
    }

    private static (byte[] Pdf, List<OoxPdfDiagnostic> Diagnostics) ConvertPictureTest(string path, OoxPdfDocxMarkupMode mode)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        using var input = File.OpenRead(path);
        using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
            FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true),
            InputKind = OoxPdfInputKind.Docx,
            DocxMarkupMode = mode,
            DocxMarkupGeometryMode = mode == OoxPdfDocxMarkupMode.AllMarkup
                ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,
            FixedCreationDate = DateTimeOffset.UnixEpoch,
            DiagnosticSink = diagnostics.Add
        });
        return (output.ToArray(), diagnostics);
    }

    public static void DocxMissingPicturePreservesIndependentControlFrames()
    {
        foreach (var mode in Modes)
        {
            var present = Render(TestFixtures.PublicCasePath("ImageRecovery/picture-present.docx"), mode);
            foreach (string variant in new[] { "missing", "corrupt" })
            {
                var actual = Render(TestFixtures.PublicCasePath($"ImageRecovery/picture-{variant}.docx"), mode);
                AssertSameTextGeometry(present.Pages, actual.Pages, "Inline image", "After picture");
                TestAssert.Equal(0, actual.Pages.Sum(page => page.Images.Count));
                TestAssert.Equal(1, actual.Diagnostics.Count);
                var diagnostic = actual.Diagnostics.Single();
                TestAssert.Equal(variant == "missing" ? "IMAGE_MISSING_PART" : "IMAGE_UNSUPPORTED_FORMAT", diagnostic.Id);
                TestAssert.Equal(OoxPdfSeverity.Error, diagnostic.Severity);
                TestAssert.Equal("/word/media/image1.png", diagnostic.PartName);
                TestAssert.Equal(1, diagnostic.PageIndex!.Value);
            }
            TestAssert.Equal(0, present.Diagnostics.Count);
            TestAssert.Equal(1, present.Pages.Sum(page => page.Images.Count));
        }
    }

    public static void DocxMissingPicturePreservesCellAndStaticStoryGeometry()
    {
        foreach (var mode in Modes)
        foreach (var (source, linkedCell) in new[]
        {
            ("docx-static-image-header.docx", false),
            ("docx-static-image-footer.docx", false),
            ("CompositionProbes/justified-linked-image-cell.docx", true)
        })
        {
            string input = TestFixtures.PublicCasePath(source);
            var document = DocxTests.ReadDocx(input, mode);
            var present = Render(document, mode);
            var pictureTargets = DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .SelectMany(p => p.Images).Select(image => image.ClickTarget).Where(target => target is not null).ToHashSet();
            TestAssert.True(present.Pages.Sum(page => page.Images.Count) > 0, "Control must paint its picture.");
            Mutate(input, parts => {
                foreach (string name in parts.Keys.Where(name => name.StartsWith("word/media/") && name.EndsWith(".png")).ToArray())
                    parts.Remove(name);
            }, path => {
                var missing = Render(path, mode);
                AssertSameTextGeometry(present.Pages, missing.Pages, linkedCell
                    ? ["Observed series.", "calculation", "Following prose."]
                    : ["Before", "After", "Following body."]);
                TestAssert.Equal(0, missing.Pages.Sum(page => page.Images.Count));
                TestAssert.Equal(1, missing.Diagnostics.Count(d => d.Id == "IMAGE_MISSING_PART"));
                var survivingLinks = missing.Pages.SelectMany(page => page.Annotations).ToArray();
                var controlLinks = present.Pages.SelectMany(page => page.Annotations)
                    .Where(link => !pictureTargets.Contains(link.Uri));
                TestAssert.True(controlLinks.Select(link => (link.Uri, link.Destination))
                    .SequenceEqual(survivingLinks.Select(link => (link.Uri, link.Destination))),
                    "Picture recovery must retain the neighboring link targets.");
                if (linkedCell)
                {
                    TestAssert.True(survivingLinks.Any(link => link.Destination.HasValue && link.Width > 0 && link.Height > 0),
                        "The authored cell hyperlink must retain a measurable destination after picture recovery.");
                }
            });
        }
    }

    public static void DocxMissingPicturePreservesDanglingAndVmlFrames()
    {
        string source = TestFixtures.PublicCasePath("ImageRecovery/picture-present.docx");
        AssertMissingPartKeepsFrame(source, parts => {
            const string relsPart = "word/_rels/document.xml.rels";
            var rels = XDocument.Parse(Encoding.UTF8.GetString(parts[relsPart]));
            rels.Root!.Elements().Single(e => (string?)e.Attribute("Id") == "rId4").Remove();
            parts[relsPart] = Encoding.UTF8.GetBytes(rels.ToString());
        });
        Mutate(source, parts => {
            XNamespace v = "urn:schemas-microsoft-com:vml";
            const string documentPart = "word/document.xml";
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[documentPart]));
            xml.Descendants(W + "drawing").Single().ReplaceWith(new XElement(W + "pict",
                new XElement(v + "shape", new XAttribute("style", "width:72pt;height:36pt"),
                    new XElement(v + "imagedata", new XAttribute(R + "id", "rId4")))));
            parts[documentPart] = Encoding.UTF8.GetBytes(xml.ToString());
        }, vml => AssertMissingPartKeepsFrame(vml, parts => parts.Remove("word/media/image1.png")));

        static void AssertMissingPartKeepsFrame(string source, Action<Dictionary<string, byte[]>> change)
        {
            var present = Render(source, OoxPdfDocxMarkupMode.Final);
            Mutate(source, change, path => {
                var missing = Render(path, OoxPdfDocxMarkupMode.Final);
                AssertSameTextGeometry(present.Pages, missing.Pages, "Inline image", "After picture");
                TestAssert.Equal(1, missing.Diagnostics.Count(d => d.Id == "IMAGE_MISSING_PART"));
                TestAssert.Equal(0, missing.Pages.Sum(page => page.Images.Count));
            });
        }
    }

    public static void DocxMissingPictureReusesFailureAndOmitsItsLinks()
    {
        Mutate(TestFixtures.PublicCasePath("ImageRecovery/picture-missing.docx"), parts => {
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
            var picture = xml.Descendants(WP+"inline").Single();
            picture.Element(WP+"docPr")!.Add(new XElement(A+"hlinkClick", new XAttribute(R+"id", "pictureLink")));
            var paragraph = picture.Ancestors(W+"p").First();
            paragraph.AddAfterSelf(new XElement(paragraph));
            parts["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
            var rels = XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
            rels.Root!.Add(new XElement(rels.Root.Name.Namespace+"Relationship", new XAttribute("Id", "pictureLink"),
                new XAttribute("Type", R.NamespaceName+"/hyperlink"), new XAttribute("TargetMode", "External"),
                new XAttribute("Target", "https://example.org/picture")));
            parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString());
        }, path => {
            var result = Render(path, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal(1, result.Diagnostics.Count(d => d.Id == "IMAGE_MISSING_PART"));
            TestAssert.Equal(0, result.Pages.Sum(page => page.Annotations.Count));
            TestAssert.Equal(0, result.Pages.Sum(page => page.Images.Count));
            TestAssert.Contains("After picture", DocxPageFootnoteNumberingTests.DecodeText(result.Pages.Single()));
        });
    }

    private static (PdfPage[] Pages, List<OoxPdfDiagnostic> Diagnostics) Render(string path, OoxPdfDocxMarkupMode mode) =>
        Render(DocxTests.ReadDocx(path, mode), mode);

    private static (PdfPage[] Pages, List<OoxPdfDiagnostic> Diagnostics) Render(DocxDocument document, OoxPdfDocxMarkupMode mode)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = new DocxRenderer(new TestFaceFontResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
        return (pages, diagnostics);
    }

    private static void AssertSameTextGeometry(PdfPage[] expected, PdfPage[] actual, params string[] authoredText)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        var text = new string[expected.Length];
        for (int i = 0; i < expected.Length; i++)
        {
            text[i] = DocxPageFootnoteNumberingTests.DecodeText(actual[i]);
            TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(expected[i]), text[i]);
            string Matrices(PdfPage page) => string.Join("\n", Regex.Matches(page.Content, @"[-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ Tm").Select(m => m.Value));
            TestAssert.True(Matrices(expected[i]).Length > 0, "Control must contain positioned text.");
            TestAssert.Equal(Matrices(expected[i]), Matrices(actual[i]));
        }
        string documentText = string.Join("\n", text);
        foreach (string phrase in authoredText) TestAssert.Contains(phrase, documentText);
    }

    private static void Mutate(string source, Action<Dictionary<string, byte[]>> change, Action<string> check)
    {
        using var archive = ZipFile.OpenRead(source);
        var parts = archive.Entries.ToDictionary(e => e.FullName, e => { using var stream=e.Open(); using var buffer=new MemoryStream(); stream.CopyTo(buffer); return buffer.ToArray(); });
        change(parts);
        string path = TestFixtures.WriteTempPackage(".docx", parts);
        try { check(path); } finally { File.Delete(path); }
    }
}
