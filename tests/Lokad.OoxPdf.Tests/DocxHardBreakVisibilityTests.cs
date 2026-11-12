using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxHardBreakVisibilityTests
{
    public static void DocxHiddenHardBreaksPreserveComposedPagesAndPaint()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string fixture in new[] { "inline-math-navigation", "footnote-image-math" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var baseline = renderer.RenderPagesAsync(Read(null, "direct"), null, default).ToBlockingEnumerable().ToArray();
            foreach (string kind in new[] { "page", "column" })
            foreach (string carrier in new[] { "direct", "sdt", "hyperlink" })
            {
                var document = Read(kind, carrier);
                TestAssert.True(!document.BodyElements.Any(element => element is DocxPageBreakElement or DocxManualBreakElement),
                    "A hidden run cannot manufacture a body boundary: " + kind + "/" + carrier);
                var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    TestAssert.Equal(baseline[index].Content, pages[index].Content);
                    TestAssert.True(baseline[index].Annotations.SequenceEqual(pages[index].Annotations),
                        "Hidden breaks preserve source links and note destinations.");
                    TestAssert.Equal(baseline[index].Images.Count, pages[index].Images.Count);
                }
            }

            DocxDocument Read(string? kind, string carrier) => DocxInlineCompositionTests.ReadMutation(
                "CompositionProbes/" + fixture, mode, false, false, (xml, w, m) =>
                {
                    if (kind is null) return;
                    var paragraph = xml.Root!.Element(w + "body")!.Elements(w + "p")
                        .Single(p => p.Elements(m + "oMath").Count() == 2);
                    var run = new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")),
                        new XElement(w + "br", new XAttribute(w + "type", kind)));
                    XElement item = carrier switch
                    {
                        "sdt" => new(w + "sdt", new XElement(w + "sdtContent", run)),
                        "hyperlink" => new(w + "hyperlink", new XAttribute(w + "anchor", "HiddenBoundary"), run),
                        _ => run
                    };
                    paragraph.Elements(m + "oMath").Last().AddBeforeSelf(item);
                });
        }
    }

    public static void DocxHardBreaksUseInheritedVisibilityAndKeepVisibleOverrides()
    {
        foreach (string owner in new[] { "direct", "character", "paragraph", "default-paragraph", "defaults", "table" })
        foreach (string kind in new[] { "page", "column" })
        foreach (bool visible in new[] { false, true })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            string runProperties = owner == "direct" ? "<w:vanish/>" :
                owner == "character" ? "<w:rStyle w:val='HiddenDerived'/>" : "";
            if (visible) runProperties = (owner == "direct" ? "" : runProperties) + "<w:vanish w:val='off'/>";
            string style = owner is "character" or "paragraph" or "default-paragraph"
                ? $"<w:style w:type='{(owner == "character" ? "character" : "paragraph")}' w:styleId='HiddenBase'><w:rPr><w:vanish/></w:rPr></w:style>" +
                    $"<w:style w:type='{(owner == "character" ? "character" : "paragraph")}' w:styleId='HiddenDerived'" +
                    (owner == "default-paragraph" ? " w:default='1'" : "") + "><w:basedOn w:val='HiddenBase'/></w:style>"
                : owner == "defaults" ? "<w:docDefaults><w:rPrDefault><w:rPr><w:vanish/></w:rPr></w:rPrDefault></w:docDefaults>"
                : owner == "table" ? "<w:style w:type='table' w:styleId='HiddenTable'><w:rPr><w:vanish/></w:rPr></w:style>" : "";
            string paragraph = "<w:p>" + (owner == "paragraph" ? "<w:pPr><w:pStyle w:val='HiddenDerived'/></w:pPr>" : "") +
                $"<w:r><w:rPr>{runProperties}</w:rPr><w:br w:type='{kind}'/></w:r></w:p>";
            string content = owner == "table" ? "<w:tbl><w:tblPr><w:tblStyle w:val='HiddenTable'/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w='4000'/></w:tblGrid><w:tr><w:tc>" + paragraph + "</w:tc></w:tr></w:tbl>" : paragraph;
            using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
                    "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>" +
                    "<Default Extension='xml' ContentType='application/xml'/>" +
                    "<Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/></Types>",
                ["_rels/.rels"] = "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                    "<Relationship Id='main' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>",
                ["word/_rels/document.xml.rels"] = "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
                    "<Relationship Id='styles' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/></Relationships>",
                ["word/styles.xml"] = "<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" + style + "</w:styles>",
                ["word/document.xml"] = "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body>" +
                    content + "<w:sectPr/></w:body></w:document>"
            });
            var document = new DocxReader().Read(OoxPackage.Open(input, default), null, default, markupMode: mode);
            var elements = owner == "table" ? document.Tables.Single().Rows.Single().Cells.Single().BodyElements : document.BodyElements;
            int breaks = elements.Count(element => element is DocxPageBreakElement or DocxManualBreakElement);
            int expected = visible && (owner != "table" || kind == "page") ? 1 : 0;
            TestAssert.True(expected == breaks,
                "Resolved visibility controls structural admission; cell columns retain their transparent fallback: " + owner + "/" + kind + "/" + mode);
        }
    }
}
