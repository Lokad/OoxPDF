using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxContainerTableGridsAlignWholeTablesInBodyAndHeader()
    {
        // The fixture is 11906 twips wide with two 1440-twip margins and a
        // 3600-twip table. The established left-indent control anchors the
        // first text rather than the grid, so include its 72-twip cell inset.
        foreach (var (fixture, part, alignment, indent) in new[]
        {
            ("body-zero", "word/document.xml", "center", 2785),
            ("header-zero", "word/static-inline-ordinary-header-box.xml", "right", 5498)
        })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Source(string? value, int leftIndent) => ReadMutation("TableGridOrigins/Containers/" + fixture,
                mode, false, false, (xml, w, _) =>
                {
                    XElement properties = xml.Descendants(w + "tblPr").Single();
                    properties.Element(w + "tblInd")!.SetAttributeValue(w + "w", leftIndent);
                    if (value is not null) properties.Element(w + "tblW")!.AddAfterSelf(new XElement(w + "jc", new XAttribute(w + "val", value)));
                }, part);

            var renderer = MathContentControlRenderer(mode);
            var expected = renderer.RenderPagesAsync(Source(null, indent), null, default).ToBlockingEnumerable().ToArray();
            var actual = renderer.RenderPagesAsync(Source(alignment, 480), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, expected.Length);
            TestAssert.True(expected[0].Annotations.Any(link => link.Uri == "https://example.invalid/static-box-table"),
                "The offset control must retain its authored link.");
            AssertMathContentControlPages(expected, actual);
        }
    }

    public static void DocxContainerTableGridsResolveAlignmentFromTableStyles()
    {
        const string fixture = "TableGridOrigins/Containers/body-zero";
        DocxTable Read(string? direct, bool mixedRows = false, bool rtl = false) => ReadMutation(fixture,
            OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
            {
                XElement properties = xml.Descendants(w + "tblPr").Single();
                properties.AddFirst(new XElement(w + "tblStyle", new XAttribute(w + "val", "DerivedAlignment")));
                if (direct is not null) properties.Element(w + "tblW")!.AddAfterSelf(new XElement(w + "jc", new XAttribute(w + "val", direct)));
                if (mixedRows)
                {
                    XElement row = xml.Descendants(w + "tr").First();
                    row.Element(w + "tc")!.AddBeforeSelf(new XElement(w + "trPr", new XElement(w + "jc", new XAttribute(w + "val", "left"))));
                }
            }, packageChange: zip =>
            {
                var entry = zip.GetEntry("word/styles.xml")!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                XElement root = xml.Root!;
                var w = root.Name.Namespace;
                root.Add(new XElement(w + "style", new XAttribute(w + "type", "table"), new XAttribute(w + "styleId", "BaseAlignment"),
                    new XElement(w + "tblPr", rtl ? new XElement(w + "bidiVisual") : null,
                        new XElement(w + "jc", new XAttribute(w + "val", "center")))),
                    new XElement(w + "style", new XAttribute(w + "type", "table"), new XAttribute(w + "styleId", "DerivedAlignment"),
                        new XElement(w + "basedOn", new XAttribute(w + "val", "BaseAlignment"))));
                entry.Delete();
                using var output = zip.CreateEntry("word/styles.xml").Open();
                xml.Save(output);
            }).Tables.Single();

        TestAssert.Equal(DocxTextAlignment.Center, Read(null).Alignment);
        TestAssert.Equal(DocxTextAlignment.Right, Read("right").Alignment);
        TestAssert.Equal(DocxTextAlignment.Left, Read("left").Alignment);
        TestAssert.Equal(DocxTextAlignment.Left, Read(null, mixedRows: true).Alignment);
        TestAssert.Equal(DocxTextAlignment.Left, Read(null, rtl: true).Alignment);
    }
}
