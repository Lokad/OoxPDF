using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxContextualLeadingFixture
{
    internal static readonly XNamespace Word = OoxNamespaces.WordprocessingNamespace;

    internal static void AddStyle(XDocument styles, string id, bool isDefault = false)
    {
        var style = new XElement(Word + "style", new XAttribute(Word + "type", "paragraph"),
            new XAttribute(Word + "styleId", id), new XElement(Word + "name", new XAttribute(Word + "val", id)));
        if (isDefault) style.SetAttributeValue(Word + "default", "1");
        styles.Root!.Add(style);
    }

    internal static void Near(double expected, double actual, string message) =>
        TestAssert.True(Math.Abs(expected - actual) < .00001d, message);

    internal static XElement[] Body(XDocument document) =>
        document.Root!.Element(Word + "body")!.Elements(Word + "p").ToArray();

    internal static XElement Spacing(XElement paragraph) =>
        paragraph.Element(Word + "pPr")!.Element(Word + "spacing")!;

    internal static DocxTextLineLayout[] Lines(DocxLayout layout) =>
        layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().ToArray();

    internal static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode mode) =>
        new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, mode);

    internal static DocxDocument ReadContextual(bool notes, Action<XDocument, XDocument>? modify = null, bool removeStyles = false)
    {
        string name = notes ? "note-marked-minimum-line-20.docx" : "ordinary-minimum-line-20.docx";
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name));
        var parts = new Dictionary<string, string>();
        using (ZipArchive zip = ZipFile.OpenRead(input))
        {
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                parts[entry.FullName] = reader.ReadToEnd();
            }
        }

        var document = XDocument.Parse(parts["word/document.xml"]);
        var styles = XDocument.Parse(parts["word/styles.xml"]);
        foreach (XElement paragraph in Body(document))
        {
            XElement spacing = Spacing(paragraph);
            spacing.SetAttributeValue(Word + "before", "80");
            spacing.SetAttributeValue(Word + "after", "80");
            paragraph.Element(Word + "pPr")!.Add(new XElement(Word + "contextualSpacing"));
        }

        modify?.Invoke(document, styles);
        parts["word/document.xml"] = document.ToString();
        if (removeStyles)
        {
            parts.Remove("word/styles.xml");
        }
        else
        {
            parts["word/styles.xml"] = styles.ToString();
        }

        if (notes)
        {
            var footnotes = XDocument.Parse(parts["word/footnotes.xml"]);
            foreach (XElement paragraph in footnotes.Descendants(Word + "p"))
            {
                XElement? properties = paragraph.Element(Word + "pPr");
                if (properties is null)
                {
                    properties = new XElement(Word + "pPr");
                    paragraph.AddFirst(properties);
                }

                XElement? spacing = properties.Element(Word + "spacing");
                if (spacing is null)
                {
                    spacing = new XElement(Word + "spacing");
                    properties.Add(spacing);
                }

                spacing.SetAttributeValue(Word + "before", "0");
                properties.Add(new XElement(Word + "contextualSpacing", new XAttribute(Word + "val", "0")));
            }

            parts["word/footnotes.xml"] = footnotes.ToString();
        }

        using MemoryStream stream = TestFixtures.CreateZipPackage(parts);
        return DocxTests.ReadDocx(stream, OoxPdfDocxMarkupMode.AllMarkup);
    }

    internal static DocxDocument ReadElevenPoint(bool notes, Action<XDocument, XDocument>? modify = null, bool removeStyles = false) =>
        ReadContextual(notes, (document, styles) =>
        {
            foreach (XElement paragraph in Body(document))
            {
                foreach (XElement run in paragraph.Elements(Word + "r"))
                {
                    XElement? properties = run.Element(Word + "rPr");
                    if (properties is null)
                    {
                        properties = new XElement(Word + "rPr");
                        run.AddFirst(properties);
                    }

                    XElement? size = properties.Element(Word + "sz");
                    if (size is null)
                    {
                        size = new XElement(Word + "sz");
                        properties.Add(size);
                    }

                    size.SetAttributeValue(Word + "val", "22");
                    if (run.Element(Word + "t") is { } text)
                    {
                        text.Value += " AB.";
                    }
                }
            }

            modify?.Invoke(document, styles);
        }, removeStyles);
}
