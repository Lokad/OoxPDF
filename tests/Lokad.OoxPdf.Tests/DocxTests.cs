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

internal static class DocxTests
{
    public static void DocxSyntheticDocumentProducesOnePdfPage()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p/>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("<< /Type /Pages /Count 1 /Kids [3 0 R] >>", pdf);
        TestAssert.Contains("/MediaBox [0 0 612 792]", pdf);
    }

    public static void DocxSyntheticA4PageSizeUsesWordPdfMediaBox()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p/><w:sectPr><w:pgSz w:w="11900" w:h="16840"/></w:sectPr></w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/MediaBox [0 0 594.96 842.04]", pdf);
    }

    public static void DocxReaderPreservesPageSettingTokens()
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
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p/>
                    <w:sectPr>
                      <w:pgSz w:w="16840" w:h="11900" w:orient="landscape"/>
                      <w:pgMar w:top="720" w:right="1440" w:bottom="1080" w:left="1800"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxPageSettings settings = document.PageSettings;

        TestAssert.Equal("16840", settings.WidthValue ?? string.Empty);
        TestAssert.Equal("11900", settings.HeightValue ?? string.Empty);
        TestAssert.Equal("landscape", settings.OrientationValue ?? string.Empty);
        TestAssert.Equal("720", settings.MarginTopValue ?? string.Empty);
        TestAssert.Equal("1440", settings.MarginRightValue ?? string.Empty);
        TestAssert.Equal("1080", settings.MarginBottomValue ?? string.Empty);
        TestAssert.Equal("1800", settings.MarginLeftValue ?? string.Empty);
        TestAssert.Equal(842d, document.PageWidthPoints);
        TestAssert.Equal(595d, document.PageHeightPoints);
        TestAssert.Equal(90d, document.MarginLeftPoints);
    }

    public static void DocxSyntheticParagraphRendersText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:jc w:val="center"/></w:pPr>
                      <w:r>
                        <w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/><w:b/><w:u w:val="single"/></w:rPr>
                        <w:t>Hello DOCX</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Contains("/F1 14 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("> Tj", pdf);
        TestAssert.Contains(" l S", pdf);
    }

    public static void DocxSyntheticStylesApplyToParagraphText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:docDefaults>
                    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/><w:color w:val="222222"/></w:rPr></w:rPrDefault>
                    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300"/></w:pPr></w:pPrDefault>
                  </w:docDefaults>
                  <w:style w:type="paragraph" w:styleId="Heading">
                    <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="240" w:line="360"/></w:pPr>
                    <w:rPr><w:sz w:val="36"/><w:b/></w:rPr>
                  </w:style>
                  <w:style w:type="character" w:styleId="Emphasis">
                    <w:rPr><w:color w:val="0000FF"/><w:i/><w:u w:val="single"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Heading"/></w:pPr>
                      <w:r>
                        <w:rPr><w:rStyle w:val="Emphasis"/></w:rPr>
                        <w:t>Styled DOCX</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 18 Tf", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains("0.213", pdf);
        TestAssert.Contains(" l S", pdf);
    }

    public static void DocxReaderParsesOnOffRunProperties()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:b w:val="on"/><w:i w:val="off"/></w:rPr><w:t>OnOff</w:t></w:r>
                      <w:r><w:rPr><w:b w:val="off"/><w:i w:val="on"/></w:rPr><w:t>OffOn</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxParagraph paragraph = document.Paragraphs[0];
        TestAssert.True(paragraph.Runs[0].Bold, "Expected w:b w:val=\"on\" to enable bold.");
        TestAssert.True(paragraph.Runs[0].Italic == false, "Expected w:i w:val=\"off\" to disable italic.");
        TestAssert.True(paragraph.Runs[1].Bold == false, "Expected w:b w:val=\"off\" to disable bold.");
        TestAssert.True(paragraph.Runs[1].Italic, "Expected w:i w:val=\"on\" to enable italic.");
    }

    public static void DocxReaderPreservesParagraphAlignmentTokens()
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
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="Justified">
                    <w:pPr><w:jc w:val="both"/></w:pPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:jc w:val="center"/></w:pPr>
                      <w:r><w:t>Center</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Justified"/></w:pPr>
                      <w:r><w:t>Inherited both</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:jc w:val="distribute"/></w:pPr>
                      <w:r><w:t>Distributed</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Default</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.Equal(DocxTextAlignment.Center, document.Paragraphs[0].Alignment);
        TestAssert.Equal("center", document.Paragraphs[0].AlignmentValue ?? string.Empty);
        TestAssert.Equal(DocxTextAlignment.Left, document.Paragraphs[1].Alignment);
        TestAssert.Equal("both", document.Paragraphs[1].AlignmentValue ?? string.Empty);
        TestAssert.Equal(DocxTextAlignment.Left, document.Paragraphs[2].Alignment);
        TestAssert.Equal("distribute", document.Paragraphs[2].AlignmentValue ?? string.Empty);
        TestAssert.True(document.Paragraphs[3].AlignmentValue is null, "Expected default alignment to keep a null source token.");
    }

    public static void DocxReaderPreservesParagraphSpacingAndKeepTokens()
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:docDefaults>
                    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300"/></w:pPr></w:pPrDefault>
                  </w:docDefaults>
                  <w:style w:type="paragraph" w:styleId="Risky">
                    <w:pPr>
                      <w:keepNext/>
                      <w:keepLines w:val="0"/>
                      <w:widowControl w:val="1"/>
                      <w:contextualSpacing/>
                      <w:spacing w:beforeAutospacing="1" w:afterLines="240" w:lineRule="exact"/>
                    </w:pPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="Risky"/>
                        <w:spacing w:before="360" w:afterAutospacing="1" w:line="480"/>
                      </w:pPr>
                      <w:r><w:t>Styled spacing</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("Risky", paragraph.StyleId ?? string.Empty);
        TestAssert.Equal(18d, paragraph.SpacingBeforePoints);
        TestAssert.Equal(6d, paragraph.SpacingAfterPoints);
        TestAssert.Equal("360", paragraph.Spacing.BeforeValue ?? string.Empty);
        TestAssert.Equal("120", paragraph.Spacing.AfterValue ?? string.Empty);
        TestAssert.Equal("240", paragraph.Spacing.AfterLinesValue ?? string.Empty);
        TestAssert.Equal("1", paragraph.Spacing.BeforeAutoSpacingValue ?? string.Empty);
        TestAssert.Equal("1", paragraph.Spacing.AfterAutoSpacingValue ?? string.Empty);
        TestAssert.Equal("480", paragraph.Spacing.LineValue ?? string.Empty);
        TestAssert.Equal("exact", paragraph.Spacing.LineRuleValue ?? string.Empty);
        TestAssert.True(paragraph.Spacing.ContextualSpacing == true, "Style contextual spacing should survive the paragraph cascade.");
        TestAssert.True(paragraph.KeepRules.KeepNext == true, "Style keepNext should survive the paragraph cascade.");
        TestAssert.True(paragraph.KeepRules.KeepLines == false, "Explicit off keepLines should survive the paragraph cascade.");
        TestAssert.True(paragraph.KeepRules.WidowControl == true, "Style widowControl should survive the paragraph cascade.");
    }

    public static void DocxReaderAppliesParagraphLineBasedSpacing()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:spacing w:beforeLines="150" w:afterLines="200"/></w:pPr>
                      <w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>Line spacing</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal(37.5d, paragraph.SpacingBeforePoints);
        TestAssert.Equal(50d, paragraph.SpacingAfterPoints);
        TestAssert.Equal("150", paragraph.Spacing.BeforeLinesValue ?? string.Empty);
        TestAssert.Equal("200", paragraph.Spacing.AfterLinesValue ?? string.Empty);
    }

    public static void DocxReaderAppliesAutoLineSpacingAsTwoHundredFortieths()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr>
                      <w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>Auto line spacing</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal(1.15d, paragraph.LineSpacingFactor);
    }

    public static void DocxReaderPreservesFontTableAlternatesAndThemeFonts()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
                  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
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
                  <Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
                  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Font metadata</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/fontTable.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:font w:name="Corporate Sans">
                    <w:altName w:val="Aptos"/>
                    <w:panose1 w:val="020B0604020202020204"/>
                    <w:family w:val="swiss"/>
                    <w:pitch w:val="variable"/>
                  </w:font>
                </w:fonts>
                """,
            ["word/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
                  <a:themeElements>
                    <a:fontScheme name="Office">
                      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
                      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxFontTableEntry entry = document.FontCatalog.Entries.Single();
        TestAssert.Equal("Corporate Sans", entry.Name);
        TestAssert.Equal("Aptos", entry.AlternateName ?? string.Empty);
        TestAssert.Equal("swiss", entry.FamilyValue ?? string.Empty);
        TestAssert.Equal("variable", entry.PitchValue ?? string.Empty);
        TestAssert.Equal("020B0604020202020204", entry.PanoseValue ?? string.Empty);
        TestAssert.Equal("Aptos Display", document.FontCatalog.ThemeFonts.MajorLatinTypeface ?? string.Empty);
        TestAssert.Equal("Aptos", document.FontCatalog.ThemeFonts.MinorLatinTypeface ?? string.Empty);
    }

    public static void DocxReaderPreservesRunFontTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:rPr>
                          <w:rFonts w:ascii="Corporate Sans" w:hAnsi="Corporate Sans" w:eastAsia="Yu Gothic" w:cs="Arial"
                            w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:csTheme="minorBidi"/>
                        </w:rPr>
                        <w:t>Font tokens</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxRunFonts fonts = document.Paragraphs.Single().Runs.Single().Fonts;
        TestAssert.Equal("Corporate Sans", fonts.Ascii ?? string.Empty);
        TestAssert.Equal("Corporate Sans", fonts.HighAnsi ?? string.Empty);
        TestAssert.Equal("Yu Gothic", fonts.EastAsia ?? string.Empty);
        TestAssert.Equal("Arial", fonts.ComplexScript ?? string.Empty);
        TestAssert.Equal("minorHAnsi", fonts.AsciiTheme ?? string.Empty);
        TestAssert.Equal("minorHAnsi", fonts.HighAnsiTheme ?? string.Empty);
        TestAssert.Equal("minorEastAsia", fonts.EastAsiaTheme ?? string.Empty);
        TestAssert.Equal("minorBidi", fonts.ComplexScriptTheme ?? string.Empty);
    }

    public static void DocxReaderCascadesRunFontTokensFromStyles()
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:docDefaults>
                    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Default Sans" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:rPrDefault>
                  </w:docDefaults>
                  <w:style w:type="paragraph" w:styleId="Body">
                    <w:rPr><w:rFonts w:ascii="Paragraph Sans" w:eastAsia="Paragraph East"/></w:rPr>
                  </w:style>
                  <w:style w:type="character" w:styleId="Emphasis">
                    <w:rPr><w:rFonts w:hAnsi="Character Sans" w:csTheme="majorBidi"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Body"/></w:pPr>
                      <w:r>
                        <w:rPr><w:rStyle w:val="Emphasis"/><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr>
                        <w:t>Styled font tokens</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTextRun run = document.Paragraphs.Single().Runs.Single();
        TestAssert.Equal("Paragraph Sans", run.FontFamily ?? string.Empty);
        TestAssert.Equal("Paragraph Sans", run.Fonts.Ascii ?? string.Empty);
        TestAssert.Equal("Character Sans", run.Fonts.HighAnsi ?? string.Empty);
        TestAssert.Equal("Paragraph East", run.Fonts.EastAsia ?? string.Empty);
        TestAssert.Equal("majorHAnsi", run.Fonts.AsciiTheme ?? string.Empty);
        TestAssert.Equal("minorHAnsi", run.Fonts.HighAnsiTheme ?? string.Empty);
        TestAssert.Equal("majorBidi", run.Fonts.ComplexScriptTheme ?? string.Empty);
    }

    public static void DocxReaderCascadesRunFontTokensThroughBasedOnStyles()
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="Base">
                    <w:rPr><w:rFonts w:ascii="Base Sans" w:hAnsi="Base Sans"/></w:rPr>
                  </w:style>
                  <w:style w:type="paragraph" w:styleId="Child">
                    <w:basedOn w:val="Base"/>
                    <w:rPr><w:rFonts w:eastAsia="Child East"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Child"/></w:pPr>
                      <w:r><w:t>Inherited style font</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTextRun run = document.Paragraphs.Single().Runs.Single();
        TestAssert.Equal("Base Sans", run.FontFamily ?? string.Empty);
        TestAssert.Equal("Base Sans", run.Fonts.Ascii ?? string.Empty);
        TestAssert.Equal("Base Sans", run.Fonts.HighAnsi ?? string.Empty);
        TestAssert.Equal("Child East", run.Fonts.EastAsia ?? string.Empty);
    }

    public static void DocxFontResolverBuildsLatinTypefaceCandidatesFromCatalogAndTheme()
    {
        var catalog = new DocxFontCatalog(
            [new DocxFontTableEntry("Corporate Sans", "Aptos", "swiss", "variable", null)],
            new DocxThemeFonts("Aptos Display", "Aptos"));
        var run = new DocxTextRun("Text", 11d, null, false, false, false, null, "Corporate Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Corporate Sans",
                HighAnsi: null,
                EastAsia: null,
                ComplexScript: null,
                AsciiTheme: "minorHAnsi",
                HighAnsiTheme: null,
                EastAsiaTheme: null,
                ComplexScriptTheme: null)
        };

        DocxTypefaceCandidates candidates = DocxFontResolver.ResolveLatinTypeface(run, catalog);

        TestAssert.Equal("Corporate Sans", candidates.Primary ?? string.Empty);
        TestAssert.Equal("Aptos", candidates.Alternate ?? string.Empty);
        TestAssert.Equal("Aptos", candidates.Theme ?? string.Empty);
    }

    public static void DocxFontPlanResolvesRunTypefaceFromFontTableAlternate()
    {
        var run = new DocxTextRun("Text", 11d, null, false, false, false, null, "Corporate Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Corporate Sans",
                HighAnsi: null,
                EastAsia: null,
                ComplexScript: null,
                AsciiTheme: null,
                HighAnsiTheme: "minorHAnsi",
                EastAsiaTheme: null,
                ComplexScriptTheme: null)
        };
        DocxDocument document = CreateFontPlanDocument(
            run,
            new DocxFontCatalog(
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null)],
                new DocxThemeFonts("Theme Display", "Theme Sans")));
        var resolver = new MapFontResolver(["Installed Sans", "Theme Sans"], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver).Runs.Single();

        TestAssert.Equal(DocxTypefaceResolutionSource.FontTableAlternate, resolved.Source);
        TestAssert.Equal("Installed Sans", resolved.RequestedFamily ?? string.Empty);
        TestAssert.Equal("Installed Sans", resolved.ResolvedFamily ?? string.Empty);
        TestAssert.Equal("Corporate Sans|Installed Sans|Theme Sans", string.Join("|", resolved.CandidateFamilies));
    }

    public static void DocxFontPlanKeepsPrimaryBeforeAlternateAndTheme()
    {
        var run = new DocxTextRun("Text", 11d, null, true, true, false, null, "Corporate Sans")
        {
            Fonts = new DocxRunFonts(
                Ascii: "Corporate Sans",
                HighAnsi: null,
                EastAsia: null,
                ComplexScript: null,
                AsciiTheme: "majorHAnsi",
                HighAnsiTheme: null,
                EastAsiaTheme: null,
                ComplexScriptTheme: null)
        };
        DocxDocument document = CreateFontPlanDocument(
            run,
            new DocxFontCatalog(
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null)],
                new DocxThemeFonts("Theme Display", "Theme Sans")));
        var resolver = new MapFontResolver(["Corporate Sans", "Installed Sans", "Theme Display"], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver).Runs.Single();

        TestAssert.Equal(DocxTypefaceResolutionSource.Primary, resolved.Source);
        TestAssert.Equal("Corporate Sans", resolved.RequestedFamily ?? string.Empty);
        TestAssert.Equal("Corporate Sans", resolved.ResolvedFamily ?? string.Empty);
        TestAssert.True(resolved.Resolution?.Bold == true && resolved.Resolution?.Italic == true, "Expected run style to flow into the font request.");
    }

    public static void DocxFontPlanIncludesPlainTableCellText()
    {
        DocxTable table = CreateSingleCellTable("Cell text", 12d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        var resolver = new MapFontResolver(["Body Sans"], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver).Runs.Single();

        TestAssert.Equal("Cell text", resolved.Run.Text);
        TestAssert.Equal(DocxTypefaceResolutionSource.Missing, resolved.Source);
        TestAssert.Equal(0, resolved.CandidateFamilies.Count);
    }

    public static void DocxFontPlanSnapshotReportsPrivateSafeCounts()
    {
        var primaryRun = new DocxTextRun("Primary", 11d, null, false, false, false, null, "Primary Sans")
        {
            Fonts = new DocxRunFonts("Primary Sans", null, null, null, null, null, null, null)
        };
        var alternateRun = new DocxTextRun("Alternate", 11d, null, false, false, false, null, "Corporate Sans")
        {
            Fonts = new DocxRunFonts("Corporate Sans", null, null, null, null, null, null, null)
        };
        var missingRun = new DocxTextRun("Missing", 11d, null, false, false, false, null, null);
        DocxDocument document = CreateFontPlanDocument(
            [primaryRun, alternateRun, missingRun],
            new DocxFontCatalog(
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null)],
                DocxThemeFonts.Empty));
        var resolver = new MapFontResolver(["Primary Sans", "Installed Sans"], "Resolver Fallback");

        DocxFontPlanSnapshot snapshot = new DocxRenderer(resolver).InspectFontPlan(document);

        TestAssert.Equal(3, snapshot.RunCount);
        TestAssert.Equal(1, snapshot.PrimaryCount);
        TestAssert.Equal(1, snapshot.FontTableAlternateCount);
        TestAssert.Equal(0, snapshot.ThemeCount);
        TestAssert.Equal(0, snapshot.ResolverFallbackCount);
        TestAssert.Equal(1, snapshot.MissingCount);
        TestAssert.Equal(3, snapshot.DistinctCandidateFamilyCount);
        TestAssert.Equal(2, snapshot.DistinctResolvedFamilyCount);
    }

    public static void DocxFontPlanTextMeasurerUsesResolvedFontFace()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        string text = "Office metrics";
        var run = new DocxTextRun(text, 12d, null, font.Value.Resolution.Bold, font.Value.Resolution.Italic, false, null, font.Value.Resolution.FamilyName)
        {
            Fonts = new DocxRunFonts(font.Value.Resolution.FamilyName, null, null, null, null, null, null, null)
        };
        DocxDocument document = CreateFontPlanDocument(run, new DocxFontCatalog([], DocxThemeFonts.Empty));
        var resolver = new SingleResolutionFontResolver(font.Value.Resolution);
        DocxFontPlan plan = DocxFontPlan.Create(document, resolver);

        double measured = new DocxFontPlanTextMeasurer(plan).MeasureText(run, text, run.FontSize);
        double expected = MeasureOpenTypeText(font.Value.Font, text, run.FontSize);

        TestAssert.True(Math.Abs(measured - expected) < 0.000001d, "Font-plan measurement should use the resolved OpenType face, including TTC face index, instead of a hard-coded font.");
    }

    public static void DocxLayoutStagePositionsMixedRunSegmentsWithRunAwareMeasurer()
    {
        var narrowRun = new DocxTextRun("A", 12d, null, false, false, false, null, "Narrow");
        var wideRun = new DocxTextRun("B", 12d, null, false, false, false, null, "Wide");
        var paragraph = new DocxParagraph(
            [narrowRun, wideRun],
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
            null);
        DocxDocument document = CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal(5d, line.Segments[0].Width);
        TestAssert.Equal(40d, line.Segments[1].Width);
        TestAssert.Equal(line.Segments[0].X + line.Segments[0].Width, line.Segments[1].X);
    }

    public static void DocxLayoutStageWrapsMixedRunTextWithRunAwareWidths()
    {
        var narrowRun = new DocxTextRun("A", 12d, null, false, false, false, null, "Narrow");
        var wideRun = new DocxTextRun(" B", 12d, null, false, false, false, null, "Wide");
        var paragraph = new DocxParagraph(
            [narrowRun, wideRun],
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
            null);
        var document = new DocxDocument(
            60d,
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("A ", lines[0].Text);
        TestAssert.Equal("B", lines[1].Text);
        TestAssert.Equal("Wide", lines[1].Segments.Single().StyleRun.FontFamily ?? string.Empty);
    }

    public static void DocxRendererUsesThemeTypefaceForThemeOnlyDefaultRun()
    {
        var resolver = new WindowsFontResolver(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"));
        string defaultFamily = resolver.Resolve(new FontRequest(DocxRenderer.DefaultDocumentTypefaceRequest)).FamilyName;
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFontExcept(defaultFamily);
        if (font is null)
        {
            return;
        }

        string family = System.Security.SecurityElement.Escape(font.Value.Resolution.FamilyName) ?? font.Value.Resolution.FamilyName;
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
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
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:docDefaults>
                    <w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:rPrDefault>
                  </w:docDefaults>
                </w:styles>
                """,
            ["word/theme/theme1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme">
                  <a:themeElements>
                    <a:fontScheme name="Theme Fonts">
                      <a:majorFont><a:latin typeface="{{family}}"/></a:majorFont>
                      <a:minorFont><a:latin typeface="{{family}}"/></a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Theme default</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/BaseFont /" + PdfEmbeddedFont.SanitizeName("LOKAD+" + font.Value.Resolution.FamilyName + "-"), pdf);
    }

    public static void DocxRendererUsesFontCatalogAlternateBeforeResolverFallback()
    {
        var resolver = new WindowsFontResolver(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"));
        string defaultFamily = resolver.Resolve(new FontRequest(DocxRenderer.DefaultDocumentTypefaceRequest)).FamilyName;
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFontExcept(defaultFamily);
        if (font is null)
        {
            return;
        }

        string family = System.Security.SecurityElement.Escape(font.Value.Resolution.FamilyName) ?? font.Value.Resolution.FamilyName;
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
                  <Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
                </Relationships>
                """,
            ["word/fontTable.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:font w:name="Unavailable Corporate Face">
                    <w:altName w:val="{{family}}"/>
                  </w:font>
                </w:fonts>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:rPr><w:rFonts w:ascii="Unavailable Corporate Face" w:hAnsi="Unavailable Corporate Face"/></w:rPr><w:t>Alternate font</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/BaseFont /" + PdfEmbeddedFont.SanitizeName("LOKAD+" + font.Value.Resolution.FamilyName + "-"), pdf);
    }

    public static void DocxRendererEmbedsResolvedTrueTypeCollectionFace()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        (FontResolution Resolution, string FirstFamily)? collectionFace = resolver.GetDiscoveredFonts()
            .Where(f => f.FontFaceIndex > 0 && f.FontFilePath is not null && Regex.IsMatch(f.FamilyName, @"^[A-Za-z0-9 ._-]+$"))
            .Select(f => TryLoadCollectionFace(f))
            .Where(item => item is not null)
            .Select(item => item!.Value)
            .FirstOrDefault(item => !item.Resolution.FamilyName.Equals(item.FirstFamily, StringComparison.OrdinalIgnoreCase) &&
                !PdfEmbeddedFont.SanitizeName("LOKAD+" + item.Resolution.FamilyName + "-").StartsWith(PdfEmbeddedFont.SanitizeName("LOKAD+" + item.FirstFamily + "-"), StringComparison.Ordinal) &&
                !PdfEmbeddedFont.SanitizeName("LOKAD+" + item.FirstFamily + "-").StartsWith(PdfEmbeddedFont.SanitizeName("LOKAD+" + item.Resolution.FamilyName + "-"), StringComparison.Ordinal));
        if (collectionFace is null)
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:rPr><w:rFonts w:ascii="{{collectionFace.Value.Resolution.FamilyName}}" w:hAnsi="{{collectionFace.Value.Resolution.FamilyName}}"/></w:rPr>
                        <w:t>Collection face</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string expected = PdfEmbeddedFont.SanitizeName("LOKAD+" + collectionFace.Value.Resolution.FamilyName + "-");
        string firstFace = PdfEmbeddedFont.SanitizeName("LOKAD+" + collectionFace.Value.FirstFamily + "-");
        TestAssert.Contains("/BaseFont /" + expected, pdf);
        TestAssert.True(!pdf.Contains("/BaseFont /" + firstFace, StringComparison.Ordinal), "Expected DOCX embedding to honor the resolved TrueType collection face index.");
    }

    public static void DocxReaderPreservesParagraphRunUnderlineTokens()
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
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="character" w:styleId="WaveUnderline">
                    <w:rPr><w:u w:val="wave"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Single</w:t></w:r>
                      <w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>None</w:t></w:r>
                      <w:r><w:rPr><w:rStyle w:val="WaveUnderline"/></w:rPr><w:t>Wave</w:t></w:r>
                      <w:r><w:t>Default</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxTextRun[] runs = document.Paragraphs[0].Runs.ToArray();

        TestAssert.True(runs[0].Underline, "Expected w:u single to keep underline enabled.");
        TestAssert.Equal("single", runs[0].UnderlineValue ?? string.Empty);
        TestAssert.True(!runs[1].Underline, "Expected w:u none to disable underline.");
        TestAssert.Equal("none", runs[1].UnderlineValue ?? string.Empty);
        TestAssert.True(runs[2].Underline, "Expected inherited w:u wave to keep underline enabled.");
        TestAssert.Equal("wave", runs[2].UnderlineValue ?? string.Empty);
        TestAssert.True(runs[3].UnderlineValue is null, "Expected missing underline to keep a null source token.");
    }

    public static void DocxParagraphLayoutPreservesSoftLineBreaks()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Alpha</w:t><w:br/><w:t>Beta</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "AlphaBeta".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal("Alpha\nBeta", document.Paragraphs[0].Runs[0].Text);
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("Alpha", lines[0].Text);
        TestAssert.Equal("Beta", lines[1].Text);
        TestAssert.True(lines[1].BaselineY < lines[0].BaselineY, "Soft line break should advance to a lower baseline.");
    }

    public static void DocxParagraphLayoutPreservesAuthoredSpaces()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun(" Alpha  Beta ", 11d, null, false, false, false, null, null)],
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
            null);
        DocxDocument document = CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), " Alpha Beta".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(" Alpha  Beta ", line.Text);
        TestAssert.True(line.Width > embedded.MeasureTextPoints("Alpha Beta", 11d), "Preserved spaces should contribute to layout width.");
    }

    public static void DocxSyntheticParagraphsBreakAcrossPages()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var body = new StringBuilder();
        for (int i = 0; i < 45; i++)
        {
            body.AppendLine($"""
                    <w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Paragraph {i}</w:t></w:r></w:p>
                """);
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                {{body}}
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Type /Pages /Count 2", pdf);
    }

    public static void DocxSyntheticPageBreakBeforeStartsNewPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>First</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Type /Pages /Count 2", pdf);
    }

    public static void DocxReaderPageBreakBeforePreservesOnOffToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>First</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr><w:r><w:t>No break</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pageBreakBefore w:val="on"/></w:pPr><w:r><w:t>Break on</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Break implicit</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxPageBreakElement[] breaks = document.BodyElements.OfType<DocxPageBreakElement>().ToArray();

        TestAssert.Equal(2, breaks.Length);
        TestAssert.Equal("pageBreakBefore", breaks[0].SourceKind);
        TestAssert.Equal("on", breaks[0].Value ?? string.Empty);
        TestAssert.Equal("pageBreakBefore", breaks[1].SourceKind);
        TestAssert.True(breaks[1].Value is null, "Expected implicit pageBreakBefore to keep a null source token.");
    }

    public static void DocxReaderPreservesParagraphSectionBreakTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:type w:val="continuous"/>
                          <w:pgSz w:w="10080" w:h="12240" w:orient="portrait"/>
                          <w:pgMar w:top="360" w:right="720" w:bottom="1080" w:left="1440"/>
                          <w:cols w:num="2" w:equalWidth="0" w:space="720"/>
                        </w:sectPr>
                      </w:pPr>
                      <w:r><w:t>Section end</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxSectionBreakElement sectionBreak = document.BodyElements.OfType<DocxSectionBreakElement>().Single();

        TestAssert.Equal("continuous", sectionBreak.TypeValue ?? string.Empty);
        TestAssert.Equal("10080", sectionBreak.PageSettings.WidthValue ?? string.Empty);
        TestAssert.Equal("12240", sectionBreak.PageSettings.HeightValue ?? string.Empty);
        TestAssert.Equal("portrait", sectionBreak.PageSettings.OrientationValue ?? string.Empty);
        TestAssert.Equal("360", sectionBreak.PageSettings.MarginTopValue ?? string.Empty);
        TestAssert.Equal("720", sectionBreak.PageSettings.MarginRightValue ?? string.Empty);
        TestAssert.Equal("1080", sectionBreak.PageSettings.MarginBottomValue ?? string.Empty);
        TestAssert.Equal("1440", sectionBreak.PageSettings.MarginLeftValue ?? string.Empty);
        TestAssert.Equal("2", sectionBreak.ColumnCountValue ?? string.Empty);
        TestAssert.Equal("0", sectionBreak.ColumnEqualWidthValue ?? string.Empty);
        TestAssert.Equal("720", sectionBreak.ColumnSpaceValue ?? string.Empty);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Paragraph section break should remain anchored after its paragraph.");
        TestAssert.True(document.BodyElements[1] is DocxSectionBreakElement, "Section break should be part of body flow.");
    }

    public static void DocxSyntheticExactLineHeightPositionsNextParagraph()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:spacing w:after="0" w:line="720" w:lineRule="exact"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
                    <w:p><w:pPr><w:spacing w:after="0" w:line="720" w:lineRule="exact"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] baselines = ExtractTextBaselines(pdf);
        TestAssert.True(baselines.Length >= 2, "Expected at least two rendered text baselines.");
        TestAssert.True(Math.Abs((baselines[0] - baselines[1]) - 36d) < 0.01d, "Exact DOCX line height should advance the next paragraph by 36 points.");
    }

    public static void DocxSyntheticContextualSpacingSuppressesSameStyleGap()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var spacing = new DocxParagraphSpacing(null, null, null, null, null, null, null, null, true);
        var first = new DocxParagraph(
            [new DocxTextRun("First", 10d, null, false, false, false, null, null)],
            [],
            "Body",
            DocxTextAlignment.Left,
            null,
            10d,
            10d,
            1d,
            10d,
            spacing,
            DocxParagraphKeepRules.Empty,
            null);
        var second = first with { Runs = [new DocxTextRun("Second", 10d, null, false, false, false, null, null)] };
        var document = new DocxDocument(
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
            [new DocxParagraphElement(first), new DocxParagraphElement(second)],
            [first, second],
            []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "FirstSecond".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(10d, Math.Round(lines[0].BaselineY - lines[1].BaselineY, 3));
    }

    public static void DocxSyntheticParagraphKeepLinesStartsBlockOnNextPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        DocxParagraph kept = CreateDocxLayoutParagraph(
            "First Second",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(null, null, true, null, null, null));
        DocxParagraph[] fillers =
        [
            CreateDocxLayoutParagraph("Fill", fontSize: 15d, lineSpacingPoints: 15d),
            CreateDocxLayoutParagraph("Fill", fontSize: 15d, lineSpacingPoints: 15d),
            CreateDocxLayoutParagraph("Fill", fontSize: 15d, lineSpacingPoints: 15d)
        ];
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Append(new DocxParagraphElement(kept))
            .ToArray();
        var document = new DocxDocument(
            54d,
            80d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            body,
            body.OfType<DocxParagraphElement>().Select(element => element.Paragraph).ToArray(),
            []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "FillFirstSecond".EnumerateRunes().Select(rune => rune.Value));

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(3, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(2, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Count());
    }

    public static void DocxSyntheticParagraphKeepNextMovesPairToNextPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        DocxParagraph keepNext = CreateDocxLayoutParagraph(
            "Keep",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph next = CreateDocxLayoutParagraph("Next", fontSize: 11d, lineSpacingPoints: 11d);
        DocxParagraph[] fillers =
        [
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
        ];
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Concat([new DocxParagraphElement(keepNext), new DocxParagraphElement(next)])
            .ToArray();
        var document = new DocxDocument(
            160d,
            80d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            body,
            body.OfType<DocxParagraphElement>().Select(element => element.Paragraph).ToArray(),
            []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "FillKeepNext".EnumerateRunes().Select(rune => rune.Value));

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(4, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(2, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Count());
    }

    public static void DocxSyntheticParagraphKeepNextChainsAcrossConsecutiveParagraphsToNextPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        DocxParagraph keepFirst = CreateDocxLayoutParagraph(
            "Keep",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph keepSecond = CreateDocxLayoutParagraph(
            "Chain",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph end = CreateDocxLayoutParagraph("End", fontSize: 11d, lineSpacingPoints: 11d);
        DocxParagraph[] fillers =
        [
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 9d, lineSpacingPoints: 9d),
            CreateDocxLayoutParagraph("Fill", fontSize: 9d, lineSpacingPoints: 9d)
        ];
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Concat([
                new DocxParagraphElement(keepFirst),
                new DocxParagraphElement(keepSecond),
                new DocxParagraphElement(end)
            ])
            .ToArray();
        var document = new DocxDocument(
            160d,
            80d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            body,
            body.OfType<DocxParagraphElement>().Select(element => element.Paragraph).ToArray(),
            []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "FillKeepChainEnd".EnumerateRunes().Select(rune => rune.Value));

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded);
        DocxTextLineLayout[] secondPageLines = layout.Pages[1].Items.OfType<DocxTextLineLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(4, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(3, secondPageLines.Length);
        TestAssert.Equal("Keep", secondPageLines[0].Text);
        TestAssert.Equal("Chain", secondPageLines[1].Text);
        TestAssert.Equal("End", secondPageLines[2].Text);
    }

    public static void DocxSyntheticParagraphWidowControlMovesThreeLineParagraphToNextPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        DocxParagraph widowControlled = CreateDocxLayoutParagraph(
            "One\nTwo\nThree",
            fontSize: 10d,
            lineSpacingPoints: 10d,
            keepRules: new DocxParagraphKeepRules(null, null, null, null, true, null));
        DocxParagraph[] fillers =
        [
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
        ];
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Append(new DocxParagraphElement(widowControlled))
            .ToArray();
        var document = new DocxDocument(
            160d,
            80d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            body,
            body.OfType<DocxParagraphElement>().Select(element => element.Paragraph).ToArray(),
            []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "FillOneTwoThree".EnumerateRunes().Select(rune => rune.Value));

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded);
        DocxTextLineLayout[] secondPageLines = layout.Pages[1].Items.OfType<DocxTextLineLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(4, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(3, secondPageLines.Length);
        TestAssert.Equal("One", secondPageLines[0].Text);
        TestAssert.Equal("Two", secondPageLines[1].Text);
        TestAssert.Equal("Three", secondPageLines[2].Text);
    }

    public static void DocxSyntheticNumberingRendersListLabels()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
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
        TestAssert.Equal(4, pdf.Split("> Tj", StringSplitOptions.None).Length - 1);
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
                    <w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
                    <w:lvl w:ilvl="1"><w:numFmt w:val="futureFormat"/><w:lvlText w:val="Item %2"/><w:suff w:val="space"/></w:lvl>
                    <w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val="bullet text"/><w:suff w:val="nothing"/></w:lvl>
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
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
        TestAssert.Equal("futureFormat", paragraphs[1].ListLabel?.FormatValue ?? string.Empty);
        TestAssert.Equal("Item %2", paragraphs[1].ListLabel?.LevelTextValue ?? string.Empty);
        TestAssert.Equal("space", paragraphs[1].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal("Item 1", paragraphs[1].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("bullet", paragraphs[2].ListLabel?.FormatValue ?? string.Empty);
        TestAssert.Equal("bullet text", paragraphs[2].ListLabel?.LevelTextValue ?? string.Empty);
        TestAssert.Equal("nothing", paragraphs[2].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal("\u2022", paragraphs[2].ListLabel?.Text ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.Equal("5.", document.Paragraphs[0].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("6.", document.Paragraphs[1].ListLabel?.Text ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

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
            return;
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Indented".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(72d, document.MarginLeftPoints);
        TestAssert.Equal(90d, line.X);
        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal(90d, line.Segments[0].X);
        TestAssert.Equal("Indented", line.Segments[1].Text);
        TestAssert.Equal(108d, line.Segments[1].X);
    }

    public static void DocxSyntheticNumberingSpaceSuffixPlacesTextAfterLabel()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Near".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        double expectedTextX = line.Segments[0].X + embedded.MeasureTextPoints("1. ", line.FontSize);
        TestAssert.Equal("space", document.Paragraphs[0].ListLabel?.SuffixValue ?? string.Empty);
        TestAssert.Equal(expectedTextX, line.Segments[1].X);
        TestAssert.True(line.Segments[1].X < 108d, "A space suffix should not advance text to the numbering tab stop.");
    }

    public static void DocxSyntheticNumberingWrapsContinuationLinesWithHangingWidth()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var label = new DocxListLabel(
            "1.",
            "decimal",
            "%1.",
            "space",
            "3",
            0,
            new DocxNumberingIndent(72d, null, null, 36d, "1440", null, null, "720"));
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Alpha Beta".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.True(lines.Length >= 3, "The narrow hanging continuation width should wrap continuation text more tightly than the first line.");
        TestAssert.Equal(2, lines[0].Segments.Count);
        TestAssert.True(lines[0].Segments[1].X < 108d, "The first line uses the space suffix immediately after the label.");
        for (int i = 1; i < lines.Length; i++)
        {
            TestAssert.Equal(108d, lines[i].X);
            TestAssert.True(lines[i].Width <= 36.001d, "Continuation lines should respect the hanging-indent text width.");
        }
    }

    public static void DocxSyntheticInlinePngRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="1828800" cy="914400"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/XObject", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
    }

    public static void DocxSyntheticTableCellInlinePngRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="914400" cy="914400"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [0, 255, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 1 /Height 1", pdf);
    }

    public static void DocxUnsupportedPngImageEmitsDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateUnsupportedHighBitDepthPng()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        TestAssert.True(File.Exists(output), "Unsupported DOCX image should not fail the whole conversion.");
        TestAssert.True(collector.Diagnostics.Any(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT" && d.Severity == OoxPdfSeverity.Error && d.PartName == "/word/media/image1.png"), "Unsupported DOCX image should emit a release-blocking diagnostic.");
    }

    public static void DocxReaderPreservesFloatingDrawingWrapTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="114300" distB="228600" distL="342900" distR="457200"
                                     simplePos="0" relativeHeight="251658240" behindDoc="0"
                                     locked="1" layoutInCell="1" allowOverlap="0">
                            <wp:extent cx="1828800" cy="914400"/>
                            <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:posOffset>63500</wp:posOffset></wp:positionV>
                            <wp:wrapSquare wrapText="bothSides"/>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxFloatingDrawing drawing = document.FloatingDrawings.Single();

        TestAssert.Equal("114300", drawing.DistanceTopValue ?? string.Empty);
        TestAssert.Equal("228600", drawing.DistanceBottomValue ?? string.Empty);
        TestAssert.Equal("342900", drawing.DistanceLeftValue ?? string.Empty);
        TestAssert.Equal("457200", drawing.DistanceRightValue ?? string.Empty);
        TestAssert.Equal("0", drawing.SimplePositionValue ?? string.Empty);
        TestAssert.Equal("251658240", drawing.RelativeHeightValue ?? string.Empty);
        TestAssert.Equal("0", drawing.BehindDocumentValue ?? string.Empty);
        TestAssert.Equal("1", drawing.LockedValue ?? string.Empty);
        TestAssert.Equal("1", drawing.LayoutInCellValue ?? string.Empty);
        TestAssert.Equal("0", drawing.AllowOverlapValue ?? string.Empty);
        TestAssert.Equal("1828800", drawing.ExtentCxValue ?? string.Empty);
        TestAssert.Equal("914400", drawing.ExtentCyValue ?? string.Empty);
        TestAssert.Equal("column", drawing.HorizontalRelativeFromValue ?? string.Empty);
        TestAssert.Equal("center", drawing.HorizontalAlignValue ?? string.Empty);
        TestAssert.Equal("paragraph", drawing.VerticalRelativeFromValue ?? string.Empty);
        TestAssert.Equal("63500", drawing.VerticalOffsetValue ?? string.Empty);
        TestAssert.Equal("wrapSquare", drawing.WrapKind ?? string.Empty);
        TestAssert.Equal("bothSides", drawing.WrapTextValue ?? string.Empty);
    }

    public static void DocxSyntheticTableRendersCellsAndText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:shd w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>One</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Two</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Three</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:fill="FCE5CD"/></w:tcPr><w:p><w:r><w:t>Four</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.851 0.918 0.827 rg", pdf);
        TestAssert.Contains(" re S", pdf);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Contains("> Tj", pdf);
    }

    public static void DocxReaderTablePreservesLayoutToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblLayout w:type="fixed"/><w:tblW w:w="2880" w:type="dxa"/><w:tblInd w:w="360" w:type="dxa"/><w:tblCellSpacing w:w="120" w:type="dxa"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Fixed</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblPr><w:tblLayout w:type="autofit"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>Autofit</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>Default</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.Equal("fixed", document.Tables[0].LayoutValue ?? string.Empty);
        TestAssert.Equal("2880", document.Tables[0].PreferredWidthValue ?? string.Empty);
        TestAssert.Equal("dxa", document.Tables[0].PreferredWidthType ?? string.Empty);
        TestAssert.Equal(144d, document.Tables[0].PreferredWidthPoints ?? 0d);
        TestAssert.Equal("360", document.Tables[0].IndentValue ?? string.Empty);
        TestAssert.Equal("dxa", document.Tables[0].IndentType ?? string.Empty);
        TestAssert.Equal(18d, document.Tables[0].IndentPoints ?? 0d);
        TestAssert.Equal("120", document.Tables[0].CellSpacingValue ?? string.Empty);
        TestAssert.Equal("dxa", document.Tables[0].CellSpacingType ?? string.Empty);
        TestAssert.Equal(6d, document.Tables[0].CellSpacingPoints ?? 0d);
        TestAssert.Equal("2160", document.Tables[0].Rows[0].Cells[0].PreferredWidthValue ?? string.Empty);
        TestAssert.Equal("dxa", document.Tables[0].Rows[0].Cells[0].PreferredWidthType ?? string.Empty);
        TestAssert.Equal(108d, document.Tables[0].Rows[0].Cells[0].PreferredWidthPoints ?? 0d);
        TestAssert.Equal("autofit", document.Tables[1].LayoutValue ?? string.Empty);
        TestAssert.True(document.Tables[2].LayoutValue is null, "Expected missing table layout to keep a null source token.");
    }

    public static void DocxReaderTableStyleAppliesCellShading()
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="ShadedTable">
                    <w:tblPr><w:tblCellMar><w:left w:w="240" w:type="dxa"/></w:tblCellMar></w:tblPr>
                    <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9EAD3"/></w:tcPr>
                    <w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="CFE2F3"/><w:tcBorders><w:top w:val="single" w:color="FF0000" w:sz="16"/></w:tcBorders></w:tcPr></w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="ShadedTable"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Styled</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:fill="FCE5CD"/></w:tcPr><w:p><w:r><w:t>Direct</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Base</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Base</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.Equal("ShadedTable", document.Tables[0].StyleId ?? string.Empty);
        TestAssert.Equal("CFE2F3", document.Tables[0].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("clear", document.Tables[0].Rows[0].Cells[0].ShadingValue ?? string.Empty);
        TestAssert.Equal("auto", document.Tables[0].Rows[0].Cells[0].ShadingColor ?? string.Empty);
        TestAssert.Equal("top", document.Tables[0].Rows[0].Cells[0].Borders.Single().Edge);
        TestAssert.Equal("16", document.Tables[0].Rows[0].Cells[0].Borders.Single().SizeValue ?? string.Empty);
        TestAssert.Equal("240", document.Tables[0].Rows[0].Cells[0].Margins.LeftValue ?? string.Empty);
        TestAssert.Equal(12d, document.Tables[0].Rows[0].Cells[0].Margins.LeftPoints ?? 0d);
        TestAssert.Equal("FCE5CD", document.Tables[0].Rows[0].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("D9EAD3", document.Tables[0].Rows[1].Cells[0].FillHex ?? string.Empty);
    }

    public static void DocxReaderTableStyleUsesCellConditionalFormatTokens()
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="ConditionalTable">
                    <w:tcPr><w:shd w:fill="DDDDDD"/></w:tcPr>
                    <w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="00FF00"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr></w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="ConditionalTable"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:cnfStyle w:val="001000000000" w:oddHBand="1"/></w:tcPr><w:p><w:r><w:t>Explicit band</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:cnfStyle w:val="000100000000" w:evenHBand="1"/></w:tcPr><w:p><w:r><w:t>Explicit even band</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Positional band</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Positional band</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTableCell first = document.Tables[0].Rows[0].Cells[0];
        DocxTableCell second = document.Tables[0].Rows[0].Cells[1];
        TestAssert.Equal("00FF00", first.FillHex ?? string.Empty);
        TestAssert.Equal("0000FF", second.FillHex ?? string.Empty);
        TestAssert.Equal("001000000000", first.ConditionalFormat?.Value ?? string.Empty);
        TestAssert.True(first.ConditionalFormat?.OddHorizontalBand == true, "Expected odd horizontal band conditional token to be preserved.");
        TestAssert.True(first.ConditionalFormat?.FirstRow is null, "Explicit cnfStyle should not invent first-row membership.");
        TestAssert.Equal("00FF00", document.Tables[0].Rows[1].Cells[0].FillHex ?? string.Empty);
    }

    public static void DocxReaderPreservesTableLookTokens()
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="LookTable">
                    <w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="firstCol"><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="00FF00"/></w:tcPr></w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="LookTable"/>
                        <w:tblLook w:val="0000" w:firstRow="0" w:firstColumn="1" w:noHBand="1"/>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTable table = document.Tables[0];
        TestAssert.Equal("0", table.Look?.FirstRowValue ?? string.Empty);
        TestAssert.Equal("1", table.Look?.FirstColumnValue ?? string.Empty);
        TestAssert.Equal("1", table.Look?.NoHorizontalBandValue ?? string.Empty);
        TestAssert.True(table.Look?.FirstRow == false, "Expected first-row table-look token to parse as false.");
        TestAssert.True(table.Look?.FirstColumn == true, "Expected first-column table-look token to parse as true.");
        TestAssert.True(table.Look?.NoHorizontalBand == true, "Expected no-horizontal-band table-look token to parse as true.");
    }

    public static void DocxSyntheticTableCellBordersUseAuthoredStroke()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcBorders><w:top w:val="single" w:color="FF0000" w:sz="16"/><w:bottom w:val="nil"/></w:tcBorders></w:tcPr>
                          <w:p/>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("2 w", pdf);
        TestAssert.Contains(" l S", pdf);
        TestAssert.DoesNotContain(" re S", pdf);
    }

    public static void DocxReaderTablePreservesHeaderRowToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:trPr><w:tblHeader w:val="0"/></w:trPr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.True(document.Tables[0].Rows[0].IsHeader, "Expected implicit tblHeader to mark the row as repeating.");
        TestAssert.True(document.Tables[0].Rows[0].HeaderValue is null, "Expected implicit tblHeader to preserve a null source token.");
        TestAssert.True(!document.Tables[0].Rows[1].IsHeader, "Expected tblHeader w:val=\"0\" to disable repeating.");
        TestAssert.Equal("0", document.Tables[0].Rows[1].HeaderValue ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesVerticalAlignmentToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:vAlign w:val="top"/></w:tcPr><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>Center</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Default</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.Equal("top", cells[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("center", cells[1].VerticalAlignmentValue ?? string.Empty);
        TestAssert.True(cells[2].VerticalAlignmentValue is null, "Expected missing cell vertical alignment to keep a null source token.");
    }

    public static void DocxReaderTableCellPreservesMarginTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:tcPr>
                          <w:tcMar>
                            <w:top w:w="120" w:type="dxa"/>
                            <w:right w:w="180" w:type="dxa"/>
                            <w:bottom w:w="240" w:type="dxa"/>
                            <w:left w:w="300" w:type="dxa"/>
                          </w:tcMar>
                        </w:tcPr>
                        <w:p><w:r><w:t>Margins</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxTableCellMargins margins = new DocxReader().Read(package).Tables[0].Rows[0].Cells[0].Margins;

        TestAssert.Equal("120", margins.TopValue ?? string.Empty);
        TestAssert.Equal("180", margins.RightValue ?? string.Empty);
        TestAssert.Equal("240", margins.BottomValue ?? string.Empty);
        TestAssert.Equal("300", margins.LeftValue ?? string.Empty);
        TestAssert.Equal(6d, margins.TopPoints ?? 0d);
        TestAssert.Equal(15d, margins.LeftPoints ?? 0d);
    }

    public static void DocxReaderTableCellPreservesGridSpanToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="720"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Span</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];
        TestAssert.Equal(2, cell.GridSpan);
        TestAssert.Equal("2", cell.GridSpanValue ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesBorderTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="single" w:color="112233" w:sz="12"/>
                              <w:bottom w:val="nil"/>
                              <w:start w:val="dashed" w:color="445566" w:sz="8"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p><w:r><w:t>Bordered</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        IReadOnlyList<DocxTableCellBorder> borders = document.Tables[0].Rows[0].Cells[0].Borders;

        TestAssert.Equal(3, borders.Count);
        TestAssert.Equal("top", borders[0].Edge);
        TestAssert.Equal("single", borders[0].Value ?? string.Empty);
        TestAssert.Equal("112233", borders[0].Color ?? string.Empty);
        TestAssert.Equal("12", borders[0].SizeValue ?? string.Empty);
        TestAssert.Equal("bottom", borders[1].Edge);
        TestAssert.Equal("nil", borders[1].Value ?? string.Empty);
        TestAssert.Equal("start", borders[2].Edge);
        TestAssert.Equal("dashed", borders[2].Value ?? string.Empty);
    }

    public static void DocxReaderTableBordersApplyOuterAndInsideEdges()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="111111" w:sz="8"/>
                          <w:bottom w:val="single" w:color="222222" w:sz="10"/>
                          <w:left w:val="single" w:color="333333" w:sz="12"/>
                          <w:right w:val="single" w:color="444444" w:sz="14"/>
                          <w:insideH w:val="single" w:color="555555" w:sz="16"/>
                          <w:insideV w:val="single" w:color="666666" w:sz="18"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        IReadOnlyList<DocxTableCellBorder> first = document.Tables[0].Rows[0].Cells[0].Borders;
        IReadOnlyList<DocxTableCellBorder> inner = document.Tables[0].Rows[1].Cells[1].Borders;
        TestAssert.Equal("111111", first.Single(border => border.Edge == "top").Color ?? string.Empty);
        TestAssert.Equal("333333", first.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("555555", first.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("666666", first.Single(border => border.Edge == "right").Color ?? string.Empty);
        TestAssert.Equal("222222", inner.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("444444", inner.Single(border => border.Edge == "right").Color ?? string.Empty);
    }

    public static void DocxReaderTableStyleBordersApplyOuterAndInsideEdges()
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="BorderedTable">
                    <w:tblPr>
                      <w:tblBorders>
                        <w:top w:val="single" w:color="111111" w:sz="8"/>
                        <w:bottom w:val="single" w:color="222222" w:sz="10"/>
                        <w:left w:val="single" w:color="333333" w:sz="12"/>
                        <w:right w:val="single" w:color="444444" w:sz="14"/>
                        <w:insideH w:val="single" w:color="555555" w:sz="16"/>
                        <w:insideV w:val="single" w:color="666666" w:sz="18"/>
                      </w:tblBorders>
                    </w:tblPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="BorderedTable"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        IReadOnlyList<DocxTableCellBorder> first = document.Tables[0].Rows[0].Cells[0].Borders;
        IReadOnlyList<DocxTableCellBorder> inner = document.Tables[0].Rows[1].Cells[1].Borders;
        TestAssert.Equal("111111", first.Single(border => border.Edge == "top").Color ?? string.Empty);
        TestAssert.Equal("333333", first.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("555555", first.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("666666", first.Single(border => border.Edge == "right").Color ?? string.Empty);
        TestAssert.Equal("222222", inner.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("444444", inner.Single(border => border.Edge == "right").Color ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesShadingTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:shd w:val="pct20" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>Shaded</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Default</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.Equal("D9EAD3", cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("pct20", cells[0].ShadingValue ?? string.Empty);
        TestAssert.Equal("112233", cells[0].ShadingColor ?? string.Empty);
        TestAssert.True(cells[1].ShadingValue is null, "Expected missing cell shading to keep a null source token.");
    }

    public static void DocxReaderTableCellPreservesParagraphModel()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="CellHeading">
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:rPr><w:sz w:val="32"/></w:rPr>
                  </w:style>
                  <w:style w:type="character" w:styleId="CellEmphasis">
                    <w:rPr><w:color w:val="336699"/><w:i/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
                        <w:p><w:pPr><w:pStyle w:val="CellHeading"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CellEmphasis"/><w:b/></w:rPr><w:t>Beta</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];
        TestAssert.Equal("Alpha Beta", cell.Text);
        TestAssert.Equal(2, cell.Paragraphs.Count);
        TestAssert.Equal("Alpha", cell.Paragraphs[0].Runs[0].Text);
        TestAssert.Equal(DocxTextAlignment.Center, cell.Paragraphs[1].Alignment);
        TestAssert.Equal(16d, cell.Paragraphs[1].Runs[0].FontSize);
        TestAssert.Equal("336699", cell.Paragraphs[1].Runs[0].ColorHex ?? string.Empty);
        TestAssert.True(cell.Paragraphs[1].Runs[0].Bold, "Expected table-cell paragraph runs to preserve direct run properties.");
        TestAssert.True(cell.Paragraphs[1].Runs[0].Italic, "Expected table-cell paragraph runs to preserve character styles.");
    }

    public static void DocxReaderTableCellPreservesNumberedParagraphs()
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
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];

        TestAssert.Equal("1.", cell.Paragraphs[0].ListLabel?.Text ?? string.Empty);
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Item".EnumerateRunes().Select(rune => rune.Value));
        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells[0]
            .TextLines[0];
        TestAssert.Equal("1.\tItem", line.Text);
        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal("Item", line.Segments[1].Text);
        TestAssert.True(line.Segments[1].X > line.Segments[0].X, "Numbered table-cell text should be segmented after the list label.");
    }

    public static void DocxSyntheticTableKeepsBodyOrder()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>After</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int firstText = pdf.IndexOf("> Tj", StringComparison.Ordinal);
        int tableGrid = pdf.IndexOf(" re S", StringComparison.Ordinal);
        int lastText = pdf.LastIndexOf("> Tj", StringComparison.Ordinal);
        TestAssert.True(firstText >= 0 && tableGrid > firstText && lastText > tableGrid, "DOCX tables should render in body order between surrounding paragraphs.");
    }

    public static void DocxSyntheticTableUsesRowHeights()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr><w:trPr><w:trHeight w:val="720"/></w:trPr><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 684 144 36 re S", pdf);
    }

    public static void DocxSyntheticTableRowsBreakAcrossPages()
    {
        string rows = string.Concat(Enumerable.Range(0, 8).Select(i => "<w:tr><w:tc><w:p/></w:tc></w:tr>"));
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>{{rows}}</w:tbl>
                    <w:sectPr><w:pgSz w:w="2880" w:h="2880"/><w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Type /Pages /Count 2", pdf);
    }

    public static void DocxTableLayoutStageKeepsManualPageBreakBoundary()
    {
        DocxTable first = CreateSingleCellTable("first", rowHeight: 20d);
        DocxTable second = CreateSingleCellTable("second", rowHeight: 20d);
        DocxDocument document = CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxPageBreakElement("pageBreakBefore", null),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded: null);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageRepeatsHeaderRowsAfterPageBreak()
    {
        var header = new DocxTableRow([new DocxTableCell("Header", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d, IsHeader: true);
        var first = new DocxTableRow([new DocxTableCell("First", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 50d);
        var second = new DocxTableRow([new DocxTableCell("Second", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 50d);
        var table = new DocxTable(null, [60d], [header, first, second]);
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded: null);

        TestAssert.Equal(2, layout.Pages.Count);
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        TestAssert.Equal(2, firstPageRows.Length);
        TestAssert.Equal(2, secondPageRows.Length);
        TestAssert.Equal("Header", secondPageRows[0].Cells[0].Cell.Text);
        TestAssert.Equal("Second", secondPageRows[1].Cells[0].Cell.Text);
    }

    public static void DocxTableLayoutStagePlacesCellsBeforePdfEmission()
    {
        var table = new DocxTable(
            null,
            [60d, 40d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded: null);

        DocxTableRowLayout row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        TestAssert.Equal(170d, row.Y);
        TestAssert.Equal(20d, row.Height);
        TestAssert.Equal(10d, row.Cells[0].X);
        TestAssert.Equal(60d, row.Cells[0].Width);
        TestAssert.Equal(70d, row.Cells[1].X);
        TestAssert.Equal(40d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageScalesGridToPreferredWidth()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            PreferredWidthPoints: 60d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(30d, row.Cells[0].Width);
        TestAssert.Equal(40d, row.Cells[1].X);
        TestAssert.Equal(30d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageScalesGridToPercentagePreferredWidth()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            PreferredWidthValue: "2500",
            PreferredWidthType: "pct");
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(45d, row.Cells[0].Width);
        TestAssert.Equal(55d, row.Cells[1].X);
        TestAssert.Equal(45d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageUsesFirstRowCellPreferredWidths()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty, PreferredWidthPoints: 40d),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty, PreferredWidthPoints: 80d)
            ], 20d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(40d, row.Cells[0].Width);
        TestAssert.Equal(50d, row.Cells[1].X);
        TestAssert.Equal(80d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageUsesLaterRowCellPreferredWidths()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [
                new DocxTableRow([
                    new DocxTableCell("span", [], null, null, null, null, [], DocxTableCellMargins.Empty, GridSpan: 2, GridSpanValue: "2")
                ], 20d),
                new DocxTableRow([
                    new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty, PreferredWidthPoints: 40d),
                    new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty, PreferredWidthPoints: 80d)
                ], 20d)
            ]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout[] rows = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .ToArray();

        TestAssert.Equal(2, rows.Length);
        TestAssert.Equal(120d, rows[0].Cells[0].Width);
        TestAssert.Equal(40d, rows[1].Cells[0].Width);
        TestAssert.Equal(50d, rows[1].Cells[1].X);
        TestAssert.Equal(80d, rows[1].Cells[1].Width);
    }

    public static void DocxTableLayoutStageUsesPercentageCellPreferredWidths()
    {
        var table = new DocxTable(
            null,
            [100d, 100d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty, PreferredWidthValue: "1250", PreferredWidthType: "pct"),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty, PreferredWidthValue: "3750", PreferredWidthType: "pct")
            ], 20d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(45d, row.Cells[0].Width);
        TestAssert.Equal(55d, row.Cells[1].X);
        TestAssert.Equal(135d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageAppliesGridSpanWidths()
    {
        var table = new DocxTable(
            null,
            [40d, 60d, 80d],
            [new DocxTableRow([
                new DocxTableCell("wide", [], null, null, null, null, [], DocxTableCellMargins.Empty, GridSpan: 2, GridSpanValue: "2"),
                new DocxTableCell("tail", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Cells[0].X);
        TestAssert.Equal(100d, row.Cells[0].Width);
        TestAssert.Equal(110d, row.Cells[1].X);
        TestAssert.Equal(80d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageAppliesTableIndent()
    {
        var table = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([
                new DocxTableCell("indented", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            IndentPoints: 18d,
            IndentValue: "360",
            IndentType: "dxa");
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(28d, row.Cells[0].X);
        TestAssert.Equal(40d, row.Cells[0].Width);
    }

    public static void DocxTableLayoutStageAppliesCellSpacing()
    {
        var table = new DocxTable(
            null,
            [40d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            CellSpacingPoints: 6d,
            CellSpacingValue: "120",
            CellSpacingType: "dxa");
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Cells[0].X);
        TestAssert.Equal(40d, row.Cells[0].Width);
        TestAssert.Equal(56d, row.Cells[1].X);
        TestAssert.Equal(60d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageBuildsParagraphTextLinesInsideCells()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var firstParagraph = new DocxParagraph(
            [new DocxTextRun("Alpha", 11d, null, false, false, false, null, null)],
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
            null);
        var secondParagraph = new DocxParagraph(
            [
                new DocxTextRun("B", 14d, "336699", false, true, false, null, null),
                new DocxTextRun("G", 14d, "993333", true, false, false, null, null)
            ],
            [],
            null,
            DocxTextAlignment.Center,
            "center",
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("Alpha BG", [firstParagraph, secondParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 44d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Alpha BG".EnumerateRunes().Select(rune => rune.Value));

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded);

        DocxTableCellLayout cellLayout = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
        TestAssert.Equal(2, cellLayout.TextLines.Count);
        TestAssert.Equal("Alpha", cellLayout.TextLines[0].Text);
        TestAssert.Equal(11d, cellLayout.TextLines[0].FontSize);
        TestAssert.Equal("BG", cellLayout.TextLines[1].Text);
        TestAssert.Equal(14d, cellLayout.TextLines[1].FontSize);
        TestAssert.Equal("336699", cellLayout.TextLines[1].StyleRun.ColorHex ?? string.Empty);
        TestAssert.Equal(2, cellLayout.TextLines[1].Segments.Count);
        TestAssert.Equal("B", cellLayout.TextLines[1].Segments[0].Text);
        TestAssert.Equal("336699", cellLayout.TextLines[1].Segments[0].StyleRun.ColorHex ?? string.Empty);
        TestAssert.Equal("G", cellLayout.TextLines[1].Segments[1].Text);
        TestAssert.Equal("993333", cellLayout.TextLines[1].Segments[1].StyleRun.ColorHex ?? string.Empty);
        TestAssert.True(cellLayout.TextLines[1].Segments[1].X > cellLayout.TextLines[1].Segments[0].X, "Second table-cell run segment should be positioned after the first segment.");
        TestAssert.True(cellLayout.TextLines[1].X > cellLayout.TextLines[0].X, "Centered table-cell paragraph text should be positioned from line width, not flattened at the left inset.");
        TestAssert.True(cellLayout.TextLines[1].BaselineY < cellLayout.TextLines[0].BaselineY, "Separate table-cell paragraphs should produce separate baselines.");
    }

    public static void DocxTableLayoutStageUsesCellMarginsForTextBox()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("Inset", 11d, null, false, false, false, null, null)],
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
            null);
        var margins = new DocxTableCellMargins(3d, 8d, null, 12d, "60", "160", null, "240");
        var cell = new DocxTableCell("Inset", [paragraph], null, null, null, null, [], margins);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 30d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Inset".EnumerateRunes().Select(rune => rune.Value));

        DocxTableCellLayout cellLayout = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(cellLayout.X + 12d, cellLayout.TextLines[0].X);
        TestAssert.Equal(cellLayout.Y + cellLayout.Height - 20d, cellLayout.TextLines[0].BaselineY);
    }

    public static void DocxTableLayoutStageAppliesCellVerticalAlignment()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("V", 11d, null, false, false, false, null, null)],
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
            null);
        var topCell = new DocxTableCell("V", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var centerCell = new DocxTableCell("V", [paragraph], null, null, null, "center", [], DocxTableCellMargins.Empty);
        var bottomCell = new DocxTableCell("V", [paragraph], null, null, null, "bottom", [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [40d, 40d, 40d], [new DocxTableRow([topCell, centerCell, bottomCell], 60d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "V".EnumerateRunes().Select(rune => rune.Value));

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        double topBaseline = row.Cells[0].TextLines[0].BaselineY;
        double centerBaseline = row.Cells[1].TextLines[0].BaselineY;
        double bottomBaseline = row.Cells[2].TextLines[0].BaselineY;
        TestAssert.True(centerBaseline < topBaseline, "Center-aligned cell text should move downward from the top baseline.");
        TestAssert.True(bottomBaseline < centerBaseline, "Bottom-aligned cell text should move below centered cell text.");
    }

    public static void DocxTableLayoutStageExpandsRowsToCellContent()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("First Second", 11d, null, false, false, false, null, null)],
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
            null);
        var cell = new DocxTableCell("First Second", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [34d], [new DocxTableRow([cell], 10d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "First Second".EnumerateRunes().Select(rune => rune.Value));

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(row.Height > 10d, "DOCX table rows should expand beyond a too-small declared height when cell text wraps.");
        TestAssert.True(row.Cells[0].TextLines.Count >= 2, "Expected the narrow cell to wrap content into multiple layout-owned text lines.");
    }

    public static void DocxLayoutSnapshotReportsPublicSafeCounts()
    {
        DocxTable table = CreateSingleCellTable("private text is not exposed", rowHeight: 20d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayoutSnapshot snapshot = new DocxRenderer().InspectLayout(document);

        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(1, snapshot.Pages[0].ItemCount);
        TestAssert.Equal(1, snapshot.Pages[0].TableRowCount);
        TestAssert.Equal(0, snapshot.Pages[0].TextLineCount);
        TestAssert.True(snapshot.Pages[0].VerticalUsed >= 20d, "Snapshot should report vertical consumption from laid-out table rows.");
        TestAssert.Equal(snapshot.Pages[0].VerticalUsed, snapshot.Pages[0].TableRowHeightSum);
        TestAssert.Equal(0d, snapshot.Pages[0].TextLineHeightSum);
        TestAssert.Equal(0d, snapshot.Pages[0].InlineImageHeightSum);
        DocxLayoutItemSnapshot row = snapshot.Pages[0].Items.Single();
        TestAssert.Equal("TableRow", row.Kind);
        TestAssert.Equal(1, row.CellCount);
        TestAssert.True(row.TextLength > 0, "Snapshot should expose text length only, not the text itself.");
    }

    public static void DocxSyntheticHeaderAndFooterRenderOnPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:footerReference w:type="default" r:id="rIdFooter1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Equal(3, pdf.Split("> Tj", StringSplitOptions.None).Length - 1);
    }

    public static void DocxSyntheticFooterPageFieldUsesGeneratedPageNumbers()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var body = new StringBuilder();
        for (int i = 0; i < 45; i++)
        {
            body.AppendLine($"""<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Paragraph {i}</w:t></w:r></w:p>""");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "/></w:p>
                </w:ftr>
                """,
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    {{body}}
                    <w:sectPr>
                      <w:footerReference w:type="default" r:id="rIdFooter1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>", pdf);
        TestAssert.True(pdf.Split("> Tj", StringSplitOptions.None).Length - 1 >= 47, "Footer page field should render on each generated page.");
    }

    public static void DocxUnsupportedFeaturesEmitDiagnostics()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:keepNext/>
                        <w:keepLines/>
                        <w:widowControl/>
                        <w:pageBreakBefore/>
                        <w:spacing w:line="240" w:lineRule="exact"/>
                        <w:sectPr/>
                      </w:pPr>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins><w:r><w:t>Inserted</w:t></w:r></w:ins>
                      <w:del><w:r><w:t>Deleted</w:t></w:r></w:del>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE </w:instrText></w:r>
                      <w:r><w:object/></w:r>
                      <w:r><w:drawing><wp:anchor/></w:drawing></w:r>
                      <w:r><w:footnoteReference w:id="2"/></w:r>
                      <w:r><w:endnoteReference w:id="3"/></w:r>
                      <w:r><w:br w:type="page"/></w:r>
                    </w:p>
                    <m:oMath/>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="2"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/vbaProject.bin"] = "macro"
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.Contains("DOCX_UNSUPPORTED_COMMENTS", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_ENDNOTE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_EQUATION", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_FLOATING_DRAWING", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_FOOTNOTE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_MACRO", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_MANUAL_BREAK", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_MULTI_COLUMN", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_OLE_OBJECT", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_SECTION_BREAK", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_TRACKED_CHANGES", ids);
        TestAssert.True(diagnostics.All(d => d.Severity == OoxPdfSeverity.Warning && d.PartName == "/word/document.xml"), "Unsupported DOCX diagnostics should be document-scoped warnings.");
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
                      <w:spacing w:beforeAutospacing="1"/>
                    </w:pPr>
                  </w:style>
                  <w:style w:type="table" w:styleId="StyledTable"/>
                </w:styles>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="1">
                    <w:lvl w:ilvl="0"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
                  </w:abstractNum>
                </w:numbering>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.Contains("DOCX_NUMBERING_INDENT", ids);
        TestAssert.Contains("DOCX_STYLE_PARAGRAPH_KEEP_RULE", ids);
        TestAssert.Contains("DOCX_STYLE_PARAGRAPH_SPACING", ids);
        TestAssert.Contains("DOCX_STYLE_TABLE_STYLE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_TABLE_HEADER_ROW", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_TABLE_STYLE", ids);
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_NUMBERING_INDENT" && d.PartName == "/word/numbering.xml"), "Numbering diagnostics should point to numbering.xml.");
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_STYLE_PARAGRAPH_KEEP_RULE" && d.PartName == "/word/styles.xml"), "Style diagnostics should point to styles.xml.");
    }

    private static DocxDocument CreateFontPlanDocument(DocxTextRun run, DocxFontCatalog fontCatalog)
    {
        return CreateFontPlanDocument([run], fontCatalog);
    }

    private static DocxDocument CreateFontPlanDocument(IReadOnlyList<DocxTextRun> runs, DocxFontCatalog fontCatalog)
    {
        var paragraph = new DocxParagraph(
            runs,
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
            null);
        return new DocxDocument(
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
            [])
        {
            FontCatalog = fontCatalog
        };
    }

    private static (FontResolution Resolution, string FirstFamily)? TryLoadCollectionFace(FontResolution resolution)
    {
        try
        {
            if (resolution.FontFilePath is null)
            {
                return null;
            }

            OpenTypeFont first = OpenTypeFont.Load(resolution.FontFilePath, 0);
            OpenTypeFont selected = OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
            return selected.FamilyName.Equals(resolution.FamilyName, StringComparison.OrdinalIgnoreCase)
                ? (resolution, first.FamilyName)
                : null;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    private static (FontResolution Resolution, OpenTypeFont Font)? FindUsableInstalledFont()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return null;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        foreach (FontResolution resolution in resolver.GetDiscoveredFonts().Where(font => font.FontFilePath is not null))
        {
            try
            {
                OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath!, resolution.FontFaceIndex);
                if (font.UnitsPerEm != 0 && font.MapCodePoint('A') != 0)
                {
                    return (resolution, font);
                }
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
            {
            }
        }

        return null;
    }

    private static (FontResolution Resolution, OpenTypeFont Font)? FindUsableInstalledFontExcept(string familyName)
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return null;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        foreach (FontResolution resolution in resolver.GetDiscoveredFonts().Where(font => font.FontFilePath is not null && !font.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase)))
        {
            try
            {
                OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath!, resolution.FontFaceIndex);
                if (font.UnitsPerEm != 0 && font.MapCodePoint('A') != 0)
                {
                    return (resolution, font);
                }
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
            {
            }
        }

        return null;
    }

    private static double MeasureOpenTypeText(OpenTypeFont font, string text, double fontSize)
    {
        double units = 0d;
        ushort previousGlyph = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            ushort glyph = font.MapCodePoint(rune.Value);
            if (previousGlyph != 0 && glyph != 0)
            {
                units += font.GetKerning(previousGlyph, glyph);
            }

            units += font.GetAdvanceWidth(glyph);
            previousGlyph = glyph;
        }

        return units * fontSize / font.UnitsPerEm;
    }

    private sealed class SingleResolutionFontResolver(FontResolution resolution) : IFontResolver
    {
        public FontResolution Resolve(FontRequest request)
        {
            return resolution with
            {
                Bold = request.Bold,
                Italic = request.Italic,
                IsFallback = !request.FamilyName.Equals(resolution.FamilyName, StringComparison.OrdinalIgnoreCase)
            };
        }
    }

    private sealed class FamilyWidthTextMeasurer : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            double width = run?.FontFamily == "Wide" ? 40d : 5d;
            return text.Length * width;
        }
    }

    private static DocxTable CreateSingleCellTable(string text, double rowHeight)
    {
        return new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell(text, [], null, null, null, null, [], DocxTableCellMargins.Empty)], rowHeight)]);
    }

    private static DocxDocument CreateLayoutTestDocument(IReadOnlyList<DocxBodyElement> bodyElements, IReadOnlyList<DocxTable> tables)
    {
        return new DocxDocument(
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
            bodyElements,
            [],
            tables);
    }

    private static DocxParagraph CreateDocxLayoutParagraph(
        string text,
        double fontSize,
        double lineSpacingPoints,
        DocxParagraphKeepRules? keepRules = null)
    {
        return new DocxParagraph(
            [new DocxTextRun(text, fontSize, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            lineSpacingPoints,
            DocxParagraphSpacing.Empty,
            keepRules ?? DocxParagraphKeepRules.Empty,
            null);
    }

    private static double[] ExtractTextBaselines(string pdf)
    {
        return Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .ToArray();
    }
}
