using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxTextTests
{
    public static void DocxRunStyleProvenanceKeepsEffectiveAndAuthoredOwners()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (scope, character, hasParagraph, depth, found, hasCharacter, size) in new[]
        {
            ("implicit", "Emphasis", true, 2, true, true, 20d),
            ("explicit", "Emphasis", true, 2, true, true, 20d),
            ("default", "Emphasis", true, 2, true, true, 20d),
            ("missing", "Empty", false, 1, true, false, 11d),
            ("empty", "Missing", false, 0, false, false, 11d),
        })
        {
            var document = ReadRunStyleProvenanceSource(scope, character, mode);
            var paragraphs = DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .Where(paragraph => paragraph.Runs.Any(run => run.Text.Length != 0)).ToArray();
            TestAssert.True(new[] { "BODY", "CELL", "HEADER" }.All(label =>
                paragraphs.Any(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)) == label)),
                "All three containers retain their source paragraphs, including any shared static-story aliases.");
            foreach (var paragraph in paragraphs)
            {
                TestAssert.Equal(scope == "explicit" ? "Normal" : scope == "missing" ? "Missing" : null, paragraph.StyleId);
                var run = paragraph.Runs.Single();
                var resolution = run.StyleResolution;
                TestAssert.Equal(hasParagraph, resolution.HasParagraphStyleRunProperties);
                TestAssert.Equal(depth, resolution.CharacterStyleDepth);
                TestAssert.Equal(found, resolution.CharacterStyleFound);
                TestAssert.Equal(hasCharacter, resolution.HasCharacterStyleRunProperties);
                TestAssert.True(resolution.HasDocumentDefaultRunProperties && resolution.HasDirectRunProperties,
                    "Document defaults and the explicit false remain distinct source owners.");
                TestAssert.Equal(size, run.FontSize);
                TestAssert.Equal(false, run.Bold);
                TestAssert.Equal(hasCharacter, run.Italic);
                TestAssert.Equal(hasParagraph ? "FF0000" : "000000", run.ColorHex!);
            }
        }
    }

    internal static DocxDocument ReadRunStyleProvenanceSource(string scope, string? character, OoxPdfDocxMarkupMode mode)
    {
        const string word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        string paragraphStyle = scope == "explicit" ? "<w:pStyle w:val=\"Normal\"/>" :
            scope == "missing" ? "<w:pStyle w:val=\"Missing\"/>" : "";
        string characterStyle = character is null ? "" : "<w:rStyle w:val=\"" + character + "\"/>";
        string Paragraph(string text) => "<w:p><w:pPr>" + paragraphStyle + "</w:pPr><w:r><w:rPr>" +
            characterStyle + "<w:b w:val=\"0\"/></w:rPr><w:t>" + text + "</w:t></w:r></w:p>";
        string body = Paragraph("BODY") + "<w:tbl><w:tblPr><w:tblW w:w=\"6000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr><w:tc>" +
            Paragraph("CELL") + "</w:tc></w:tr></w:tbl><w:p/>";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/><Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/></Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"main\" Type=\"{office}/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"styles\" Type=\"{office}/styles\" Target=\"styles.xml\"/><Relationship Id=\"header\" Type=\"{office}/header\" Target=\"header1.xml\"/></Relationships>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{word}\" xmlns:r=\"{office}\"><w:body>{body}<w:sectPr><w:headerReference w:type=\"default\" r:id=\"header\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\" w:header=\"360\"/></w:sectPr></w:body></w:document>",
            ["word/header1.xml"] = $"<w:hdr xmlns:w=\"{word}\">{Paragraph("HEADER")}</w:hdr>",
            ["word/styles.xml"] = $$"""
                <w:styles xmlns:w="{{word}}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="22"/><w:color w:val="000000"/></w:rPr></w:rPrDefault></w:docDefaults>
                  <w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:b/></w:rPr></w:style>
                  <w:style w:type="paragraph" w:styleId="Normal">{{(scope == "empty" ? "" : "<w:basedOn w:val=\"Base\"/><w:rPr><w:sz w:val=\"40\"/><w:color w:val=\"FF0000\"/></w:rPr>")}}</w:style>
                  {{(scope == "default" ? "<w:style w:type=\"paragraph\" w:styleId=\"BodyDefault\" w:default=\"1\"><w:basedOn w:val=\"Normal\"/></w:style>" : "")}}
                  <w:style w:type="character" w:styleId="CharBase"><w:rPr><w:i/></w:rPr></w:style>
                  <w:style w:type="character" w:styleId="Emphasis"><w:basedOn w:val="CharBase"/></w:style>
                  <w:style w:type="character" w:styleId="Empty"/>
                </w:styles>
                """,
        });
        using var package = OoxPackage.Open(input, default);
        return new DocxReader().Read(package, null, default, mode);
    }
}
