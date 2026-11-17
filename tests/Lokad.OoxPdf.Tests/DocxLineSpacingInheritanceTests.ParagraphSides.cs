using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxLineSpacingInheritanceTests
{
    public static void DocxParagraphSpacingSelectsOneEffectiveUnitAcrossOwners()
    {
        (string? Absolute, string? Lines, string? Auto, double? Expected)[] cases =
        [
            (null, null, null, null), ("0", null, "1", 14), ("4800", null, "true", 14),
            ("4800", "100", "on", 14), (null, "100", "1", 14),
            ("4800", "100", null, 15), ("4800", "100", "0", 15),
            ("4800", "0", "false", 0), ("240", null, "off", 12),
            ("0", null, "0", 0), ("240", "invalid", "0", 12)
        ];
        foreach (var sample in cases)
        foreach (string origin in new[] { "default", "style", "direct" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string attributes = Side("before") + Side("after");
            var source = ReadParagraphSideProbe(attributes, origin, mode, compose: true);
            DocxParagraph[] paragraphs =
            [
                source.BodyElements.OfType<DocxParagraphElement>().ElementAt(1).Paragraph,
                source.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single(),
                source.HeaderParagraphs.Single()
            ];
            foreach (var paragraph in paragraphs)
            {
                TestAssert.Equal(sample.Expected ?? 0d, paragraph.SpacingBeforePoints);
                TestAssert.Equal(sample.Expected ?? 8d, paragraph.SpacingAfterPoints);
                TestAssert.Equal(sample.Absolute, paragraph.Spacing.BeforeValue);
                TestAssert.Equal(sample.Absolute, paragraph.Spacing.AfterValue);
                TestAssert.Equal(sample.Lines, paragraph.Spacing.BeforeLinesValue);
                TestAssert.Equal(sample.Auto, paragraph.Spacing.AfterAutoSpacingValue);
            }
            string Side(string name) =>
                (sample.Absolute is null ? "" : $" w:{name}=\"{sample.Absolute}\"") +
                (sample.Lines is null ? "" : $" w:{name}Lines=\"{sample.Lines}\"") +
                (sample.Auto is null ? "" : $" w:{name}Autospacing=\"{sample.Auto}\"");
        }
    }

    public static void DocxParagraphAutomaticSpacingKeepsOnePageAndVisibleWords()
    {
        foreach (string side in new[] { "before", "after" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadParagraphSideProbe($"w:{side}=\"4800\" w:{side}Autospacing=\"1\"",
                "direct", mode, compose: false);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            foreach (string word in new[] { "FIRST", "MIDDLE", "LAST" })
                TestAssert.Equal(word.Length, PortableRenderingTests.ReadPaintedSourceLabel(pages[0], word).Count);
            PdfContentValidator.ValidatePage(pages[0], 0, default);
        }
    }

    private static DocxDocument ReadParagraphSideProbe(string attributes, string origin,
        OoxPdfDocxMarkupMode mode, bool compose)
    {
        const string relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string spacing = "<w:spacing " + attributes + "/>";
        string properties = origin == "direct" ? spacing : origin == "style" ? "<w:pStyle w:val=\"SideStyle\"/>" : "";
        string Paragraph(string word, bool styled) => "<w:p><w:pPr>" + (styled ? properties : "") +
            "</w:pPr><w:r><w:t>" + word + "</w:t></w:r></w:p>";
        string body = Paragraph("FIRST", false) + Paragraph("MIDDLE", true) + Paragraph("LAST", false);
        if (compose) body += "<w:tbl><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr><w:tc>" +
            Paragraph("CELL", true) + "</w:tc></w:tr></w:tbl>";
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
                (compose ? "<Override PartName=\"/word/header.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" : "") + "</Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"main\" Type=\"{office}/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"styles\" Type=\"{office}/styles\" Target=\"styles.xml\"/>" +
                (compose ? $"<Relationship Id=\"header\" Type=\"{office}/header\" Target=\"header.xml\"/>" : "") + "</Relationships>",
            ["word/styles.xml"] = $"<w:styles xmlns:w=\"{Word}\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"DejaVu Sans\" w:hAnsi=\"DejaVu Sans\"/><w:sz w:val=\"24\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line=\"300\" w:lineRule=\"auto\" {(origin == "default" ? attributes : "")}/></w:pPr></w:pPrDefault></w:docDefaults>" +
                (origin == "style" ? "<w:style w:type=\"paragraph\" w:styleId=\"SideStyle\"><w:pPr>" + spacing + "</w:pPr></w:style>" : "") + "</w:styles>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{Word}\" xmlns:r=\"{office}\"><w:body>{body}<w:sectPr>" +
                (compose ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "") +
                "<w:pgSz w:w=\"12240\" w:h=\"7200\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>"
        };
        if (compose) parts["word/header.xml"] = $"<w:hdr xmlns:w=\"{Word}\">{Paragraph("HEADER", true)}</w:hdr>";
        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        return new DocxReader().Read(package, null, default, mode);
    }
}
