using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxTableStyleInheritanceTests
{
    public static void DocxTableParagraphStylesKeepNamedAndDirectOwnership()
    {
        string paragraphs = Paragraph("Table defaults") +
            Paragraph("Named style", "<w:pStyle w:val=\"Parent\"/>") +
            Paragraph("Inherited style", "<w:pStyle w:val=\"Child\"/>") +
            Paragraph("Direct properties", "<w:pStyle w:val=\"Child\"/><w:jc w:val=\"right\"/><w:ind w:left=\"720\"/><w:spacing w:line=\"480\" w:lineRule=\"auto\" w:after=\"40\"/>") +
            Paragraph("Unknown style", "<w:pStyle w:val=\"Missing\"/>");
        string body = "<w:tbl><w:tblPr><w:tblStyle w:val=\"Table\"/><w:tblLook w:firstColumn=\"1\" w:firstRow=\"0\" w:noHBand=\"1\" w:noVBand=\"1\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr>" +
            "<w:tc>" + paragraphs + "</w:tc><w:tc>" + paragraphs + "</w:tc></w:tr></w:tbl>";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var table = ReadParagraphStyleProbe(body, mode).Tables.Single();
            for (int column = 0; column < 2; column++)
            {
                var p = table.Rows.Single().Cells[column].Paragraphs;
                TestAssert.Equal(column == 0 ? DocxTextAlignment.Center : DocxTextAlignment.Right, p[0].Alignment);
                TestAssert.Equal(72d, p[0].Indent.LeftPoints);
                TestAssert.Equal(48d, p[0].LineSpacingPoints);
                TestAssert.Equal(DocxTextAlignment.Left, p[1].Alignment);
                TestAssert.Equal(0d, p[1].Indent.LeftPoints);
                TestAssert.Equal(1d, p[1].LineSpacingFactor);
                TestAssert.True(p[1].LineSpacingPoints is null, "Named automatic spacing replaces table point units.");
                TestAssert.Equal(column == 0 ? 4d : 0d, p[1].SpacingAfterPoints);
                TestAssert.Equal(DocxTextAlignment.Left, p[2].Alignment);
                TestAssert.Equal(18d, p[2].Indent.LeftPoints);
                TestAssert.Equal(1d, p[2].LineSpacingFactor);
                TestAssert.True(p[2].LineSpacingPoints is null, "A basedOn style retains the paragraph's effective units.");
                TestAssert.Equal(6d, p[2].SpacingAfterPoints);
                TestAssert.Equal(DocxTextAlignment.Right, p[3].Alignment);
                TestAssert.Equal(36d, p[3].Indent.LeftPoints);
                TestAssert.Equal(2d, p[3].LineSpacingFactor);
                TestAssert.True(p[3].LineSpacingPoints is null, "Direct automatic spacing stays above both style levels.");
                TestAssert.Equal(2d, p[3].SpacingAfterPoints);
                TestAssert.Equal(p[0].Alignment, p[4].Alignment);
                TestAssert.Equal(p[0].LineSpacingPoints, p[4].LineSpacingPoints);
                foreach (var paragraph in p)
                {
                    TestAssert.Equal(12d, paragraph.Indent.RightPoints);
                    TestAssert.True(paragraph.StyleResolution.HasTableStyleParagraphProperties,
                        "Nonconflicting table properties and their provenance survive.");
                }
            }
        }
    }

    public static void DocxTableParagraphStylesKeepAllLabelsInsideExactHeight()
    {
        string source = "<w:p><w:pPr><w:pStyle w:val=\"Parent\"/></w:pPr>" +
            "<w:r><w:rPr><w:color w:val=\"FF0000\"/></w:rPr><w:t>ALPHA</w:t></w:r>" +
            "<w:r><w:rPr><w:color w:val=\"008000\"/></w:rPr><w:br/><w:t>BETA</w:t></w:r>" +
            "<w:r><w:rPr><w:color w:val=\"0000FF\"/></w:rPr><w:br/><w:t>GAMMA</w:t></w:r></w:p>";
        string body = "<w:tbl><w:tblPr><w:tblStyle w:val=\"Table\"/><w:tblW w:w=\"6000\" w:type=\"dxa\"/><w:tblLayout w:type=\"fixed\"/>" +
            "<w:tblLook w:firstColumn=\"0\" w:firstRow=\"0\" w:noHBand=\"1\" w:noVBand=\"1\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val=\"1200\" w:hRule=\"exact\"/></w:trPr>" +
            "<w:tc><w:tcPr><w:tcW w:w=\"6000\" w:type=\"dxa\"/></w:tcPr>" + source + "</w:tc></w:tr></w:tbl>" + Paragraph("TAIL");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadParagraphStyleProbe(body, mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(0, diagnostics.Count);
            foreach (string label in new[] { "ALPHA", "BETA", "GAMMA", "TAIL" })
                TestAssert.Equal(label.Length, PortableRenderingTests.ReadPaintedSourceLabel(page, label).Count);
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    private static string Paragraph(string text, string properties = "") =>
        "<w:p><w:pPr>" + properties + "</w:pPr><w:r><w:t>" + text + "</w:t></w:r></w:p>";

    private static DocxDocument ReadParagraphStyleProbe(string body, OoxPdfDocxMarkupMode mode)
    {
        const string word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"main\" Type=\"{office}/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"styles\" Type=\"{office}/styles\" Target=\"styles.xml\"/></Relationships>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{word}\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>",
            ["word/styles.xml"] = $$"""
                <w:styles xmlns:w="{{word}}">
                  <w:docDefaults>
                    <w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
                    <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
                  </w:docDefaults>
                  <w:style w:type="table" w:customStyle="1" w:styleId="Table">
                    <w:name w:val="Table Probe"/>
                    <w:pPr><w:jc w:val="right"/><w:ind w:left="1440" w:right="240"/><w:spacing w:line="960" w:lineRule="exact"/></w:pPr>
                    <w:tblStylePr w:type="firstCol"><w:pPr><w:jc w:val="center"/><w:spacing w:after="80"/></w:pPr></w:tblStylePr>
                  </w:style>
                  <w:style w:type="paragraph" w:styleId="Parent">
                    <w:pPr><w:jc w:val="left"/><w:ind w:left="0"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
                  </w:style>
                  <w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Parent"/>
                    <w:pPr><w:ind w:left="360"/><w:spacing w:after="120"/></w:pPr>
                  </w:style>
                </w:styles>
                """
        });
        using var package = OoxPackage.Open(input, default);
        return new DocxReader().Read(package, null, default, mode);
    }
}
