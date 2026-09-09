using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNumberingTests
{
    public static void DocxFontPlanIncludesNumberingMarkerTypeface()
    {
        var bodyRun = new DocxTextRun("Body", 11d, null, false, false, false, null, "Body Sans")
        {
            Fonts = new DocxRunFonts("Body Sans", null, null, null, null, null, null, null)
        };
        var label = new DocxListLabel(
            "#",
            "bullet",
            "#",
            "tab",
            "3",
            0,
            new DocxNumberingIndent(36d, null, null, 18d, null, "720", null, null, "360", null, null),
            new DocxTextRunStyle(null, null, null, null, null, null, "Marker Sans", new DocxRunFonts("Marker Sans", null, null, null, null, null, null, null)));
        var paragraph = new DocxParagraph(
            [bodyRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            label);
        DocxDocument document = DocxTests.CreateFontPlanDocument([paragraph], DocxFontCatalog.Empty);
        var resolver = new MapFontResolver(["Body Sans", "Marker Sans"], "Resolver Fallback");

        DocxResolvedRunTypeface marker = DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Single(run => run.Run.Text == "#");

        TestAssert.Equal("Marker Sans", marker.RequestedFamily ?? string.Empty);
        TestAssert.Equal(DocxTypefaceResolutionSource.Primary, marker.Source);
    }

    public static void DocxSyntheticNumberingRendersListLabels()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Beta</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Type0", pdf);
        // Office A/B 2026-09-07 (probe-labsize decimal): styles-less decimal labels resolve flat 12pt.
        TestAssert.Equal(2, Regex.Matches(pdf, "0\\.048 Tc").Count);
        TestAssert.Equal(8, DocxTests.CountPdfTextShows(pdf));
    }

    public static void DocxReaderMapsSymbolCharsetNumberingText()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                  <Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
                </Relationships>
                """,
            ["word/fontTable.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:font w:name="Marker Symbols"><w:charset w:val="02"/></w:font>
                </w:fonts>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0">
                      <w:numFmt w:val="bullet"/>
                      <w:lvlText w:val="§"/>
                      <w:rPr><w:rFonts w:ascii="Marker Symbols" w:hAnsi="Marker Symbols"/></w:rPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Symbol marker</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxListLabel? label = document.Paragraphs.Single().ListLabel;
        TestAssert.Equal("\uF0A7", label?.Text ?? string.Empty);
        TestAssert.Equal("Marker Symbols", label?.Style.Fonts.Ascii ?? string.Empty);
    }

    public static void DocxReaderSizesListLabelFromNumberingThenParagraphStyle()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="28"/></w:rPr></w:style>
                </w:styles>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="0">
                    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>
                  </w:abstractNum>
                  <w:abstractNum w:abstractNumId="1">
                    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/><w:sz w:val="36"/></w:rPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>
                  <w:num w:numId="8"><w:abstractNumId w:val="1"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Styled label</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Explicit label</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        // Office sizes bullets from the numbering level, then the paragraph style chain, and only then
        // direct formatting (bullet-size probe 2026-09-06).
        TestAssert.Equal(2, document.Paragraphs.Count);
        TestAssert.Equal(14d, document.Paragraphs[0].ListLabel?.Style.FontSize ?? 0d);
        TestAssert.Equal(18d, document.Paragraphs[1].ListLabel?.Style.FontSize ?? 0d);
    }

    public static void DocxReaderSizesListLabelFromDocumentDefaults()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="16"/></w:rPr></w:rPrDefault></w:docDefaults>
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                </w:styles>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="0">
                    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Small</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Large</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        // Document defaults win outright over body runs (bullet-chain probes 2026-09-06).
        TestAssert.Equal(2, document.Paragraphs.Count);
        TestAssert.Equal(8d, document.Paragraphs[0].ListLabel?.Style.FontSize ?? 0d);
        TestAssert.Equal(8d, document.Paragraphs[1].ListLabel?.Style.FontSize ?? 0d);
    }

    public static void DocxReaderCapsUnsizedListLabelAtTwelvePoints()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
                </w:styles>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="0">
                    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="36"/></w:rPr><w:t>Capped</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        // With no size anywhere in the chain, oversized body runs cap the bullet at 12pt (bullet-chain probes 2026-09-06).
        TestAssert.Equal(12d, document.Paragraphs.Single().ListLabel?.Style.FontSize ?? 0d);
    }

    public static void DocxLayoutEmitsListLabelAtResolvedSize()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="40"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>""",
            ["word/numbering.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num></w:numbering>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Sized</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxLayoutPage page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None).Pages.Single();
        DocxTextLineLayout line = page.Items.OfType<DocxTextLineLayout>().Single();
        DocxTextSegmentLayout label = line.Segments[0];
        TestAssert.Equal(DocxTextSegmentRole.ListLabel, label.Role);
        TestAssert.Equal(20d, label.FontSize ?? 0d);
    }


    public static void DocxReaderFallsBackToTwelvePointsForUnstyledCellLabels()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>""",
            ["word/numbering.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num></w:numbering>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblGrid><w:gridCol w:w="2160"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>Body</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph cellParagraph = document.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single(p => p.Runs.Any(r => r.Text == "Cell"));
        DocxParagraph bodyParagraph = document.Paragraphs.Single(p => p.Runs.Any(r => r.Text == "Body"));
        TestAssert.Equal(12d, cellParagraph.ListLabel?.Style.FontSize ?? 0d);
        // Office A/B 2026-09-07 (labsize/nosize probes): unresolvable style and defaults emit flat 12pt labels
        // at direct 9/14/18pt for family-less and Symbol lvls alike, so direct size never refines the label.
        TestAssert.Equal(12d, bodyParagraph.ListLabel?.Style.FontSize ?? 0d);
    }
    public static void DocxReaderFallsBackToTwelvePointsWithoutStyleOrDefaultSize()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:style></w:styles>""",
            ["word/numbering.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="[BULLET]"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num></w:numbering>""".Replace("[BULLET]", ((char)0xF0B7).ToString()),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="18"/></w:rPr><w:t>Nine</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="28"/></w:rPr><w:t>Fourteen</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        // Office A/B 2026-09-07 (probe-labsize nosize-styles-nodocdef): direct 9/14pt bodies emit flat 12pt labels
        // when neither the style chain nor document defaults resolve a size.
        foreach (DocxParagraph paragraph in document.Paragraphs)
        {
            TestAssert.Equal(12d, paragraph.ListLabel?.Style.FontSize ?? 0d);
        }
    }


    public static void DocxReaderPrefersDocumentDefaultsOverUnstyledCellFallback()
    {
        // Office A/B 2026-09-07 (probe-celldef): docDefaults 8pt wins over the unstyled-cell 12pt fallback
        // at direct 9/14pt bodies.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:style></w:styles>""",
            ["word/numbering.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="[BULLET]"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="0"/></w:num></w:numbering>""".Replace("[BULLET]", ((char)0xF0B7).ToString()),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblGrid><w:gridCol w:w="2160"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="18"/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph cellParagraph = document.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single(p => p.Runs.Any(r => r.Text == "Cell"));
        TestAssert.Equal(8d, cellParagraph.ListLabel?.Style.FontSize ?? 0d);
    }


    public static void DocxReaderPreservesNumberingFormatTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1)"/><w:pPr><w:ind w:left="720" w:hanging="360"/><w:tabs><w:tab w:val="num" w:pos="360"/></w:tabs></w:pPr></w:lvl>
                    <w:lvl w:ilvl="1"><w:numFmt w:val="futureFormat"/><w:lvlText w:val="Item %2"/><w:suff w:val="space"/></w:lvl>
                    <w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val="bullet text"/><w:suff w:val="nothing"/><w:rPr><w:rFonts w:ascii="Marker Sans" w:hAnsi="Marker Sans"/><w:color w:val="123456"/><w:b/></w:rPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Roman</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Future</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Bullet</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxParagraph> paragraphs = document.Paragraphs;

        TestAssert.Equal("lowerRoman", paragraphs[0].ListLabel?.FormatValue ?? string.Empty);
        TestAssert.Equal("%1)", paragraphs[0].ListLabel?.LevelTextValue ?? string.Empty);
        TestAssert.Equal("tab", paragraphs[0].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal("3)", paragraphs[0].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("3", paragraphs[0].ListLabel?.NumberId ?? string.Empty);
        TestAssert.Equal(0, paragraphs[0].ListLabel?.Level ?? -1);
        TestAssert.Equal("720", paragraphs[0].ListLabel?.Indent.LeftValue ?? string.Empty);
        TestAssert.Equal("360", paragraphs[0].ListLabel?.Indent.HangingValue ?? string.Empty);
        TestAssert.Equal(36d, paragraphs[0].ListLabel?.Indent.LeftPoints ?? 0d);
        TestAssert.Equal(18d, paragraphs[0].ListLabel?.Indent.HangingPoints ?? 0d);
        TestAssert.Equal("num", paragraphs[0].ListLabel?.Indent.NumberingTabValue ?? string.Empty);
        TestAssert.Equal("360", paragraphs[0].ListLabel?.Indent.NumberingTabPositionValue ?? string.Empty);
        TestAssert.Equal(18d, paragraphs[0].ListLabel?.Indent.NumberingTabPositionPoints ?? 0d);
        TestAssert.Equal("futureFormat", paragraphs[1].ListLabel?.FormatValue ?? string.Empty);
        TestAssert.Equal("Item %2", paragraphs[1].ListLabel?.LevelTextValue ?? string.Empty);
        TestAssert.Equal("space", paragraphs[1].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal("Item 1", paragraphs[1].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("bullet", paragraphs[2].ListLabel?.FormatValue ?? string.Empty);
        TestAssert.Equal("bullet text", paragraphs[2].ListLabel?.LevelTextValue ?? string.Empty);
        TestAssert.Equal("nothing", paragraphs[2].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal("bullet text", paragraphs[2].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("Marker Sans", paragraphs[2].ListLabel?.Style.FontFamily ?? string.Empty);
        TestAssert.Equal("Marker Sans", paragraphs[2].ListLabel?.Style.Fonts.Ascii ?? string.Empty);
        TestAssert.Equal("123456", paragraphs[2].ListLabel?.Style.ColorHex ?? string.Empty);
        TestAssert.True(paragraphs[2].ListLabel?.Style.Bold == true, "Numbering marker run properties should be preserved without font-name special cases.");
    }

    public static void DocxReaderNumberingStartOverrideRestartsList()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3">
                    <w:abstractNumId w:val="7"/>
                    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>
                  </w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Five</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Six</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal("5.", document.Paragraphs[0].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("6.", document.Paragraphs[1].ListLabel?.Text ?? string.Empty);
    }

    public static void DocxReaderNumberingLevelOverrideReplacesAbstractLevel()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="decimal"/>
                      <w:lvlText w:val="%1."/>
                      <w:suff w:val="tab"/>
                      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3">
                    <w:abstractNumId w:val="7"/>
                    <w:lvlOverride w:ilvl="0">
                      <w:startOverride w:val="9"/>
                      <w:lvl w:ilvl="0">
                        <w:start w:val="4"/>
                        <w:numFmt w:val="decimal"/>
                        <w:lvlText w:val="Item %1)"/>
                        <w:suff w:val="space"/>
                        <w:pPr><w:ind w:left="1440" w:hanging="720"/></w:pPr>
                        <w:rPr><w:rFonts w:ascii="Override Sans" w:hAnsi="Override Sans"/><w:color w:val="AA0000"/><w:b/></w:rPr>
                      </w:lvl>
                    </w:lvlOverride>
                  </w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Nine</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Ten</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxParagraph> paragraphs = document.Paragraphs;

        TestAssert.Equal("Item 9)", paragraphs[0].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("Item 10)", paragraphs[1].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("Item %1)", paragraphs[0].ListLabel?.LevelTextValue ?? string.Empty);
        TestAssert.Equal("space", paragraphs[0].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal("1440", paragraphs[0].ListLabel?.Indent.LeftValue ?? string.Empty);
        TestAssert.Equal("720", paragraphs[0].ListLabel?.Indent.HangingValue ?? string.Empty);
        TestAssert.Equal(72d, paragraphs[0].ListLabel?.Indent.LeftPoints ?? 0d);
        TestAssert.Equal(36d, paragraphs[0].ListLabel?.Indent.HangingPoints ?? 0d);
        TestAssert.Equal("Override Sans", paragraphs[0].ListLabel?.Style.Fonts.Ascii ?? string.Empty);
        TestAssert.Equal("AA0000", paragraphs[0].ListLabel?.Style.ColorHex ?? string.Empty);
        TestAssert.True(paragraphs[0].ListLabel?.Style.Bold == true, "Concrete num-level override marker style should replace the abstract numbering level.");
    }

    public static void DocxReaderNumberingResolvesMultilevelLabels()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>One one</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>One two</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Two</w:t></w:r></w:p>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Two one</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal("1.", document.Paragraphs[0].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("1.1.", document.Paragraphs[1].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("1.2.", document.Paragraphs[2].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("2.", document.Paragraphs[3].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("2.1.", document.Paragraphs[4].ListLabel?.Text ?? string.Empty);
    }

    public static void DocxSyntheticNumberingIndentMovesListLine()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="10000" w:h="4000"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Indented".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(72d, document.MarginLeftPoints);
        TestAssert.Equal(90d, line.X);
        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal(90d, line.Segments[0].X);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal("Indented", line.Segments[2].Text);
        TestAssert.Equal(108d, line.Segments[2].X);
    }

    public static void DocxSyntheticNumberingTabPositionMovesTextOnly()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/><w:tabs><w:tab w:val="num" w:pos="240"/></w:tabs></w:pPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="10000" w:h="4000"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Indented".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(90d, line.X);
        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal(90d, line.Segments[0].X);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal("Indented", line.Segments[2].Text);
        TestAssert.Equal(108d, line.Segments[2].X);
    }

    public static void DocxSyntheticNumberingTabBeyondIndentMovesTextOnly()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/><w:tabs><w:tab w:val="num" w:pos="1440"/></w:tabs></w:pPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="10000" w:h="4000"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Indented".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(90d, line.X);
        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal(90d, line.Segments[0].X);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal("Indented", line.Segments[2].Text);
        TestAssert.Equal(144d, line.Segments[2].X);
    }

    public static void DocxSyntheticNumberingParagraphIndentOverridesLevelIndent()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:ind w:left="1440" w:hanging="720"/></w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="10000" w:h="4000"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Indented".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(108d, line.X);
        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal(108d, line.Segments[0].X);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal("Indented", line.Segments[2].Text);
        TestAssert.Equal(144d, line.Segments[2].X);
    }

    public static void DocxSyntheticNumberingSpaceSuffixPlacesTextAfterLabel()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Near</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="10000" w:h="4000"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Near".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        // Office A/B 2026-09-07 (probe-labsize decimal): styles-less decimal labels resolve flat 12pt while the space
        // separator keeps the body size, so each span measures at its own segment size.
        double expectedTextX = line.Segments[0].X
            + embedded.MeasureTextPoints("1.", line.Segments[0].FontSize ?? line.FontSize)
            + embedded.MeasureTextPoints(" ", line.Segments[1].FontSize ?? line.FontSize);
        TestAssert.Equal("space", document.Paragraphs[0].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal(expectedTextX, line.Segments[2].X);
        TestAssert.True(line.Segments[2].X < 108d, "A space suffix should not advance text to the numbering tab stop.");
    }

    public static void DocxNumberingSeparatorUsesBodyRunStyle()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        var label = new DocxListLabel(
            "•",
            "bullet",
            "\uF0B7",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, "MarkerFace", new DocxRunFonts("MarkerFace", null, null, null, null, null, null, null)));
        var bodyRun = new DocxTextRun("Item", 10d, null, false, false, false, null, "BodyFace")
        {
            Fonts = new DocxRunFonts("BodyFace", null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [bodyRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            label);
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [paragraph],
            []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font.Value.Font, "• Item".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(DocxTextSegmentRole.ListLabel, line.Segments[0].Role);
        TestAssert.Equal("MarkerFace", line.Segments[0].StyleRun.EffectiveProperties.FontFamily ?? string.Empty);
        TestAssert.Equal(DocxTextSegmentRole.ListSeparator, line.Segments[1].Role);
        TestAssert.Equal("BodyFace", line.Segments[1].StyleRun.EffectiveProperties.FontFamily ?? string.Empty);
        TestAssert.Equal(DocxTextSegmentRole.Text, line.Segments[2].Role);
        TestAssert.Equal("BodyFace", line.Segments[2].StyleRun.EffectiveProperties.FontFamily ?? string.Empty);
    }

    public static void DocxSyntheticNumberingWrapsContinuationLinesWithHangingWidth()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var label = new DocxListLabel(
            "1.",
            "decimal",
            "%1.",
            "space",
            "3",
            0,
            new DocxNumberingIndent(72d, null, null, 36d, null, "1440", null, null, "720", null, null),
            DocxTextRunStyle.Empty);
        var paragraph = new DocxParagraph(
            [new DocxTextRun("Alpha Beta Alpha Beta", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            label);
        var document = new DocxDocument(
            180d,
            120d,
            36d,
            36d,
            36d,
            36d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [paragraph],
            []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Alpha Beta".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.True(lines.Length >= 3, "The narrow hanging continuation width should wrap continuation text more tightly than the first line.");
        TestAssert.Equal(3, lines[0].Segments.Count);
        TestAssert.Equal(" ", lines[0].Segments[1].Text);
        TestAssert.True(lines[0].Segments[2].X < 108d, "The first line uses the space suffix immediately after the label.");
        for (int i = 1; i < lines.Length; i++)
        {
            TestAssert.Equal(108d, lines[i].X);
            TestAssert.True(lines[i].Width <= 36.001d, "Continuation lines should respect the hanging-indent text width.");
        }
    }

    public static void DocxWordCompatibleAllMarkupPositionsNumberingInNarrowedBodyFrame()
    {
        var numberingRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "41",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["numberingChange", "numPr", "ilvl"]);
        var label = new DocxListLabel(
            "12.",
            "decimal",
            "%1.",
            "tab",
            "7",
            0,
            new DocxNumberingIndent(72d, null, null, 36d, 84d, "1440", null, null, "720", "num", "1680"),
            new DocxTextRunStyle(
                11d,
                "C00000",
                true,
                null,
                null,
                null,
                "LabelFace",
                new DocxRunFonts("LabelFace", null, null, null, null, null, null, null)));
        var bodyRun = new DocxTextRun(
            "Alpha Beta Gamma Delta Epsilon Zeta Eta Theta",
            10d,
            null,
            false,
            false,
            false,
            null,
            "BodyFace")
        {
            Fonts = new DocxRunFonts("BodyFace", null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [bodyRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            label)
        {
            Revisions = [numberingRevision],
            TabStops = [new DocxTabStop(96d, "1920", "num", null)]
        };
        DocxDocument document = new(
            360d,
            300d,
            36d,
            36d,
            36d,
            36d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [paragraph],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxLayoutPage page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .Single();
        DocxTextLineLayout[] lines = page.Items.OfType<DocxTextLineLayout>().ToArray();
        DocxTextLineLayout firstLine = lines[0];
        DocxTextSegmentLayout labelSegment = firstLine.Segments[0];
        DocxTextSegmentLayout separatorSegment = firstLine.Segments[1];
        DocxTextSegmentLayout bodySegment = firstLine.Segments[2];

        // Break-equivalent reserve (W5-R): 360/36/36 body (288) times the default print scale.
        TestAssert.True(
            Math.Abs(page.ColumnFrames.Single().Width - (288d * 0.842391d)) < 0.000000001d,
            $"Word-compatible reserve should size the layout body to the authored body times scale. Width={page.ColumnFrames.Single().Width}.");
        TestAssert.True(lines.Length > 1, "The narrowed all-markup body frame should wrap the numbered paragraph.");
        TestAssert.True(Math.Abs(firstLine.X - (36d + 36d * 0.842391d)) < 0.000001d, "First-line X should add the scaled hanging-derived label offset to the margin. X=" + firstLine.X);
        TestAssert.Equal(DocxTextSegmentRole.ListLabel, labelSegment.Role);
        TestAssert.Equal("12.", labelSegment.Text);
        TestAssert.True(Math.Abs(labelSegment.X - (36d + 36d * 0.842391d)) < 0.000001d, "Label X should match the scaled first-line start. X=" + labelSegment.X);
        TestAssert.Equal(15d, labelSegment.Width);
        TestAssert.Equal(11d, labelSegment.FontSize ?? 0d);
        TestAssert.Equal("LabelFace", labelSegment.StyleRun.EffectiveProperties.FontFamily ?? string.Empty);
        TestAssert.Equal("C00000", labelSegment.StyleRun.EffectiveProperties.ColorHex ?? string.Empty);
        TestAssert.True(labelSegment.StyleRun.EffectiveProperties.Bold, "Numbering label style should apply to the marker run.");
        TestAssert.Equal(DocxTextSegmentRole.ListSeparator, separatorSegment.Role);
        TestAssert.Equal(" ", separatorSegment.Text);
        TestAssert.True(Math.Abs(separatorSegment.X - (36d + 36d * 0.842391d + 15d)) < 0.000001d, "Separator X should follow the scaled label plus the measured label width. X=" + separatorSegment.X);
        TestAssert.Equal("BodyFace", separatorSegment.StyleRun.EffectiveProperties.FontFamily ?? string.Empty);
        TestAssert.Equal(DocxTextSegmentRole.Text, bodySegment.Role);
        TestAssert.True(Math.Abs(bodySegment.X - (36d + 96d * 0.842391d)) < 0.000001d, "Body X should add the scaled num-tab position to the margin. X=" + bodySegment.X);
        TestAssert.Equal("BodyFace", bodySegment.StyleRun.EffectiveProperties.FontFamily ?? string.Empty);
        TestAssert.True(Math.Abs(lines[1].X - (36d + 72d * 0.842391d)) < 0.000001d, "Continuation X should add the scaled hanging indent to the margin. X=" + lines[1].X);
        TestAssert.True(lines.Skip(1).All(line => Math.Abs(line.X - (36d + 72d * 0.842391d)) < 0.000001d), "Continuation lines should use the hanging indent text start.");
        TestAssert.True(
            firstLine.SourceParagraph?.Revisions.Single().PropertyElementNames.Contains("numberingChange") == true,
            "Revised numbering properties should stay attached to the numbered paragraph layout line.");
    }

    public static void DocxStyleAndNumberingLayoutRisksEmitDiagnostics()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:pStyle w:val="Risky"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="StyledTable"/></w:tblPr>
                      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="Risky">
                    <w:pPr>
                      <w:keepNext/>
                      <w:contextualSpacing/>
                      <w:spacing w:line="240" w:lineRule="mystery"/>
                    </w:pPr>
                  </w:style>
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblStylePr w:type="firstRow">
                      <w:rPr><w:bCs/></w:rPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="1">
                    <w:lvl w:ilvl="0">
                      <w:numFmt w:val="bullet"/>
                      <w:pPr><w:ind w:left="720" w:firstLineChars="120"/></w:pPr>
                      <w:rPr><w:rFonts w:ascii="Marker Sans"/></w:rPr>
                    </w:lvl>
                  </w:abstractNum>
                </w:numbering>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.Contains("DOCX_NUMBERING_INDENT", ids);
        TestAssert.Contains("DOCX_STYLE_PARAGRAPH_SPACING", ids);
        TestAssert.DoesNotContain("DOCX_STYLE_TABLE_COMPLEX_SCRIPT_RUN", ids);
        TestAssert.DoesNotContain("DOCX_NUMBERING_MARKER_FONT", ids);
        TestAssert.DoesNotContain("DOCX_STYLE_TABLE_STYLE", ids);
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_TABLE_STYLE", ids);
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_TABLE_HEADER_ROW", ids);
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_NUMBERING_INDENT" && d.PartName == "/word/numbering.xml"), "Numbering diagnostics should point to numbering.xml.");
    }

    public static void DocxSupportedNumberingLeftHangingTabsDoNotEmitIndentDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="1">
                    <w:lvl w:ilvl="0">
                      <w:numFmt w:val="decimal"/>
                      <w:pPr>
                        <w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs>
                        <w:ind w:left="720" w:hanging="360"/>
                      </w:pPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
                </w:numbering>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_NUMBERING_INDENT", ids);
        TestAssert.DoesNotContain("DOCX_NUMBERING_MARKER_FONT", ids);
    }

    public static void DocxSupportedNumberingFirstLineIndentDoesNotEmitIndentDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="1">
                    <w:lvl w:ilvl="0">
                      <w:numFmt w:val="decimal"/>
                      <w:pPr>
                        <w:tabs><w:tab w:val="num" w:pos="840"/></w:tabs>
                        <w:ind w:left="720" w:firstLine="120"/>
                      </w:pPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
                </w:numbering>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_NUMBERING_INDENT", ids);
        TestAssert.DoesNotContain("DOCX_NUMBERING_MARKER_FONT", ids);
    }
}
