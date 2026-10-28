using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxHiddenPictureTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static void DocxHiddenPicturesPreserveOfficeVisibilityAndParagraphs()
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("docx-run-picture-office.json")));
        foreach (var expected in gold.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("OfficeApp").GetString() == "Word"))
        {
            string path = Case(expected.GetProperty("File").GetString()!);
            TestAssert.Equal(expected.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
                WithDocument(Parts(path), mode, document =>
                {
                    int images = expected.GetProperty("Pages")[0].GetProperty("ImageCount").GetInt32();
                    var pages = Render(document, mode);
                    TestAssert.Equal(images, pages.Sum(page => page.Images.Count));
                    TestAssert.Equal(3, document.Paragraphs.Count);
                    TestAssert.Equal(images, document.Paragraphs[1].Images.Count);
                    TestAssert.Contains("After picture", DocxPageFootnoteNumberingTests.DecodeText(pages.Single()));
                });
        }
    }

    public static void DocxHiddenPicturesUseInheritedStylesAndExplicitFalse()
    {
        foreach (string owner in new[] { "direct", "character", "paragraph", "defaults" })
        foreach (bool visible in new[] { false, true })
        {
            var parts = Parts();
            var xml = Xml(parts, "word/document.xml");
            var run = PictureRun(xml);
            var properties = new XElement(W + "rPr"); run.AddFirst(properties);
            var styles = Xml(parts, "word/styles.xml");
            if (owner == "direct") properties.Add(new XElement(W + "vanish", new XAttribute(W + "val", visible ? "off" : "on")));
            else
            {
                var hidden = new XElement(W + "rPr", new XElement(W + "vanish"));
                if (owner == "defaults")
                {
                    var defaults = styles.Root!.Element(W + "docDefaults");
                    if (defaults is null) { defaults = new XElement(W + "docDefaults"); styles.Root.AddFirst(defaults); }
                    defaults.Element(W + "rPrDefault")?.Remove(); defaults.AddFirst(new XElement(W + "rPrDefault", hidden));
                }
                else
                {
                    styles.Root!.Add(new XElement(W + "style", new XAttribute(W + "type", owner), new XAttribute(W + "styleId", "HiddenPicture"), hidden));
                    if (owner == "character") properties.Add(new XElement(W + "rStyle", new XAttribute(W + "val", "HiddenPicture")));
                    else run.Parent!.AddFirst(new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "HiddenPicture"))));
                }
                if (visible) properties.Add(new XElement(W + "vanish", new XAttribute(W + "val", "false")));
            }
            Put(parts, "word/document.xml", xml); Put(parts, "word/styles.xml", styles);
            WithDocument(parts, OoxPdfDocxMarkupMode.Final, document =>
            {
                TestAssert.Equal(visible ? 1 : 0, document.Paragraphs[1].Images.Count);
                TestAssert.Equal(visible ? 1 : 0, Render(document, OoxPdfDocxMarkupMode.Final).Sum(page => page.Images.Count));
            });
        }
    }

    public static void DocxHiddenPicturesKeepLinksRevisionsAndContainerOwnership()
    {
        foreach (string container in new[] { "body", "cell", "header", "insertion" })
        foreach (bool hidden in new[] { false, true })
        {
            var parts = Parts(); var xml = Xml(parts, "word/document.xml"); var run = PictureRun(xml);
            if (hidden) run.AddFirst(new XElement(W + "rPr", new XElement(W + "vanish")));
            var paragraph = run.Parent!;
            run.Descendants(WP + "docPr").Single().Add(new XElement(A + "hlinkClick", new XAttribute(R + "id", "pictureLink")));
            var rels = Xml(parts, "word/_rels/document.xml.rels");
            rels.Root!.Add(new XElement(PR + "Relationship", new XAttribute("Id", "pictureLink"), new XAttribute("Type", R.NamespaceName + "/hyperlink"),
                new XAttribute("Target", "https://example.com/picture"), new XAttribute("TargetMode", "External")));
            if (container == "cell")
            {
                paragraph.ReplaceWith(new XElement(W + "tbl", new XElement(W + "tblPr"), new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "9360"))),
                    new XElement(W + "tr", new XElement(W + "tc", new XElement(W + "tcPr"), new XElement(paragraph)))));
            }
            else if (container == "header")
            {
                paragraph.Remove();
                Put(parts, "word/headerPicture.xml", new XDocument(new XElement(W + "hdr", paragraph)));
                Put(parts, "word/_rels/headerPicture.xml.rels", new XDocument(new XElement(rels.Root)));
                rels.Root.Add(new XElement(PR + "Relationship", new XAttribute("Id", "pictureHeader"), new XAttribute("Type", R.NamespaceName + "/header"), new XAttribute("Target", "headerPicture.xml")));
                xml.Descendants(W + "sectPr").Single().AddFirst(new XElement(W + "headerReference", new XAttribute(W + "type", "default"), new XAttribute(R + "id", "pictureHeader")));
                var types = Xml(parts, "[Content_Types].xml");
                types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/word/headerPicture.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
                Put(parts, "[Content_Types].xml", types);
            }
            else if (container == "insertion")
                run.ReplaceWith(new XElement(W + "ins", new XAttribute(W + "id", "42"), new XAttribute(W + "author", "Public author"), new XElement(run)));
            Put(parts, "word/document.xml", xml); Put(parts, "word/_rels/document.xml.rels", rels);
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
                WithDocument(parts, mode, document =>
                {
                    var pages = Render(document, mode);
                    TestAssert.Equal(hidden ? 0 : 1, pages.Sum(page => page.Images.Count));
                    TestAssert.Equal(hidden ? 0 : 1, pages.Sum(page => page.Annotations.Count(link => link.Uri == "https://example.com/picture")));
                });
        }
    }

    public static void DocxHiddenPicturesKeepComputedFieldAndSourceRunAdmission()
    {
        foreach (bool computed in new[] { false, true })
        {
            var parts = Parts(); var xml = Xml(parts, "word/document.xml"); var run = PictureRun(xml);
            var visible = new XElement(run);
            run.AddFirst(new XElement(W + "rPr", new XElement(W + "vanish")));
            if (computed) run.ReplaceWith(new XElement(W + "fldSimple", new XAttribute(W + "instr", "PAGE"), new XElement(run)));
            else run.AddAfterSelf(new XElement(W + "r", new XElement(W + "t", "Between pictures")), visible);
            Put(parts, "word/document.xml", xml);
            WithDocument(parts, OoxPdfDocxMarkupMode.Final, document =>
            {
                var paragraph = document.Paragraphs[1];
                TestAssert.Equal(computed ? 0 : 1, paragraph.Images.Count);
                if (!computed)
                {
                    TestAssert.Equal(2, paragraph.Images.Single().SourceRunIndex);
                    TestAssert.Equal(1, paragraph.InlineContent!.OfType<DocxInlinePicture>().Count());
                }
                TestAssert.Equal(computed ? 0 : 1, Render(document, OoxPdfDocxMarkupMode.Final).Sum(page => page.Images.Count));
            });
        }
    }

    public static void DocxHiddenPicturesAndMasterThemesCancelBeforePublication()
    {
        foreach (string name in new[] { "docx-run-picture-hidden.docx", "pptx-master-themes-distinct.pptx" })
        {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
            byte[] original = [1, 2, 3]; File.WriteAllBytes(output, original);
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            try
            {
                TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(Case(name), output,
                    new OoxPdfOptions { FontResolver = new PortableRenderingTests.PinnedResolver() }, cancelled.Token));
                TestAssert.True(original.AsSpan().SequenceEqual(File.ReadAllBytes(output)), "Cancellation preserves published output.");
                TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
            }
            finally { File.Delete(output); }
        }
    }

    private static PdfPage[] Render(DocxDocument document, OoxPdfDocxMarkupMode mode) => new DocxRenderer(
        new PortableRenderingTests.PinnedResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, default).ToArray();
    private static void WithDocument(Dictionary<string, byte[]> parts, OoxPdfDocxMarkupMode mode, Action<DocxDocument> check)
    {
        using var stream = TestFixtures.CreateZipPackage(parts);
        var package = OoxPackage.Open(stream, default);
        check(new DocxReader().Read(package, null, default, markupMode: mode));
    }
    private static XElement PictureRun(XDocument xml) => xml.Descendants(W + "r").Single(run => run.Element(W + "drawing") is not null);
    private static XDocument Xml(Dictionary<string, byte[]> parts, string name)
    {
        using var stream = new MemoryStream(parts[name]); return XDocument.Load(stream);
    }
    private static void Put(Dictionary<string, byte[]> parts, string name, XDocument xml) => parts[name] = Encoding.UTF8.GetBytes(xml.ToString());
    private static Dictionary<string, byte[]> Parts(string? path = null)
    {
        using var zip = ZipFile.OpenRead(path ?? Case("docx-run-picture-visible.docx"));
        return zip.Entries.ToDictionary(entry => entry.FullName, entry => { using var data = new MemoryStream(); using var stream = entry.Open(); stream.CopyTo(data); return data.ToArray(); });
    }
    private static string Case(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name));
}
