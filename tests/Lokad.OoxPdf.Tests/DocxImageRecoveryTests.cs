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
    private static readonly OoxPdfDocxMarkupMode[] Modes = [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup];

    public static void DocxMissingPicturePreservesIndependentControlFrames()
    {
        foreach (var mode in Modes)
        {
            var present = Render(Case("ImageRecovery/picture-present.docx"), mode);
            foreach (string variant in new[] { "missing", "corrupt" })
            {
                var actual = Render(Case("ImageRecovery/picture-"+variant+".docx"), mode);
                SameTextGeometry(present.Pages, actual.Pages);
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
        foreach (string source in new[] { "docx-static-image-header.docx", "docx-static-image-footer.docx",
            "CompositionProbes/justified-linked-image-cell.docx" })
        {
            var present = Render(Case(source), mode);
            TestAssert.True(present.Pages.Sum(page => page.Images.Count) > 0, "Control must paint its picture.");
            Mutate(Case(source), parts => {
                foreach (string name in parts.Keys.Where(name => name.StartsWith("word/media/") && name.EndsWith(".png")).ToArray())
                    parts.Remove(name);
            }, path => {
                var missing = Render(path, mode);
                SameTextGeometry(present.Pages, missing.Pages);
                TestAssert.Equal(0, missing.Pages.Sum(page => page.Images.Count));
                TestAssert.Equal(1, missing.Diagnostics.Count(d => d.Id == "IMAGE_MISSING_PART"));
                TestAssert.Equal(0, missing.Pages.Sum(page => page.Annotations.Count));
            });
        }
    }

    public static void DocxMissingPicturePreservesDanglingAndVmlFrames()
    {
        foreach (bool vml in new[] { false, true })
        {
            Action<Dictionary<string, byte[]>> change = parts => {
                var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
                if (vml)
                {
                    XNamespace v = "urn:schemas-microsoft-com:vml";
                    xml.Descendants(W+"drawing").Single().ReplaceWith(new XElement(W+"pict",
                        new XElement(v+"shape", new XAttribute("style", "width:72pt;height:36pt"),
                            new XElement(v+"imagedata", new XAttribute(R+"id", "rId4")))));
                    parts["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
                }
            };
            Mutate(Case("ImageRecovery/picture-present.docx"), change, presentPath => {
                var present = Render(presentPath, OoxPdfDocxMarkupMode.Final);
                Mutate(presentPath, parts => {
                    if (vml) parts.Remove("word/media/image1.png");
                    else
                    {
                        var rels = XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
                        rels.Root!.Elements().Where(e => (string?)e.Attribute("Id") == "rId4").Remove();
                        parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString());
                    }
                }, missingPath => {
                    var missing = Render(missingPath, OoxPdfDocxMarkupMode.Final);
                    SameTextGeometry(present.Pages, missing.Pages);
                    TestAssert.Equal(1, missing.Diagnostics.Count(d => d.Id == "IMAGE_MISSING_PART"));
                    TestAssert.Equal(0, missing.Pages.Sum(page => page.Images.Count));
                });
            });
        }
    }

    public static void DocxMissingPictureReusesFailureAndOmitsItsLinks()
    {
        Mutate(Case("ImageRecovery/picture-missing.docx"), parts => {
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

    private static (PdfPage[] Pages, List<OoxPdfDiagnostic> Diagnostics) Render(string path, OoxPdfDocxMarkupMode mode)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var document = DocxTests.ReadDocx(path, mode);
        var pages = new DocxRenderer(new TestFaceFontResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, diagnostics.Add, default).ToArray();
        return (pages, diagnostics);
    }

    private static void SameTextGeometry(PdfPage[] expected, PdfPage[] actual)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        for (int i=0; i<expected.Length; i++)
        {
            TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(expected[i]), DocxPageFootnoteNumberingTests.DecodeText(actual[i]));
            string Matrices(PdfPage page) => string.Join("\n", Regex.Matches(page.Content, @"[-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ Tm").Select(m => m.Value));
            TestAssert.True(Matrices(expected[i]).Length > 0, "Control must contain positioned text.");
            TestAssert.Equal(Matrices(expected[i]), Matrices(actual[i]));
        }
    }

    private static string Case(string path) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", path));
    private static void Mutate(string source, Action<Dictionary<string, byte[]>> change, Action<string> check)
    {
        using var archive = ZipFile.OpenRead(source);
        var parts = archive.Entries.ToDictionary(e => e.FullName, e => { using var stream=e.Open(); using var buffer=new MemoryStream(); stream.CopyTo(buffer); return buffer.ToArray(); });
        change(parts);
        string path = TestFixtures.WriteTempPackage(".docx", parts);
        try { check(path); } finally { File.Delete(path); }
    }
}
