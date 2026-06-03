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
                      <w:pgMar w:top="720" w:right="1440" w:bottom="1080" w:left="1800" w:header="1440" w:footer="1080"/>
                      <w:docGrid w:linePitch="326"/>
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
        TestAssert.Equal("1440", settings.HeaderDistanceValue ?? string.Empty);
        TestAssert.Equal("1080", settings.FooterDistanceValue ?? string.Empty);
        TestAssert.Equal(72d, settings.HeaderDistancePoints ?? 0d);
        TestAssert.Equal(54d, settings.FooterDistancePoints ?? 0d);
        TestAssert.Equal("326", settings.DocGridLinePitchValue ?? string.Empty);
        TestAssert.Equal(16.3d, settings.DocGridLinePitchPoints ?? 0d);
        TestAssert.Equal(842d, document.PageWidthPoints);
        TestAssert.Equal(595d, document.PageHeightPoints);
        TestAssert.Equal(90d, document.MarginLeftPoints);
    }

    public static void DocxReaderPreservesParagraphSnapToGridTokens()
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
                    <w:p><w:pPr><w:snapToGrid w:val="0"/></w:pPr><w:r><w:t>Free line</w:t></w:r></w:p>
                    <w:p><w:pPr><w:snapToGrid/></w:pPr><w:r><w:t>Grid line</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.True(document.Paragraphs[0].SnapToGrid == false, "Explicit w:snapToGrid val=0 should opt out.");
        TestAssert.Equal("0", document.Paragraphs[0].SnapToGridValue ?? string.Empty);
        TestAssert.True(document.Paragraphs[1].SnapToGrid == true, "Empty w:snapToGrid should opt in.");
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
        TestAssert.Contains("/F1 14.04 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(CountPdfTextShows(pdf) >= 1, "Expected DOCX paragraph text to render as a PDF text-show operation.");
        TestAssert.Contains(" re f", pdf);
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
        TestAssert.Contains(" re f", pdf);
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

    public static void DocxReaderParsesRunCharacterSpacing()
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
                      <w:r><w:rPr><w:spacing w:val="40"/></w:rPr><w:t>Wide</w:t></w:r>
                      <w:r><w:rPr><w:spacing w:val="-20"/></w:rPr><w:t>Tight</w:t></w:r>
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
        TestAssert.Equal(2d, paragraph.Runs[0].CharacterSpacingPoints);
        TestAssert.Equal(-1d, paragraph.Runs[1].CharacterSpacingPoints);
    }

    public static void DocxRendererEmitsRunCharacterSpacingAsPositionedGlyphs()
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
                      <w:r><w:rPr><w:sz w:val="24"/><w:spacing w:val="40"/></w:rPr><w:t>AB</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("TJ", pdf);
        TestAssert.Contains("-166.667", pdf);
    }

    public static void OfficePdfTextProfileExposesWordNumberedListCharacterSpacing()
    {
        TestAssert.Equal(0.048d, OfficePdfTextEmissionProfile.ObservedWordNumberedListTextStateCharacterSpacing(12d));
        TestAssert.Equal(0.044d, OfficePdfTextEmissionProfile.ObservedWordNumberedListTextStateCharacterSpacing(11d));
    }

    public static void DocxTextEmissionPlannerOwnsListLabelTextStateTarget()
    {
        var decimalLabel = new DocxListLabel("1", "decimal", "%1.", "tab", "7", 0, DocxNumberingIndent.Empty, DocxTextRunStyle.Empty);
        var bulletLabel = new DocxListLabel("*", "bullet", "\uF0B7", "tab", "7", 0, DocxNumberingIndent.Empty, DocxTextRunStyle.Empty);
        var decimalRun = new DocxTextRun("1", 12d, null, false, false, false, null, null);
        var bulletRun = new DocxTextRun("*", 12d, null, false, false, false, null, null);
        DocxTextStateCharacterSpacingTarget decimalTarget = DocxTextEmissionPlanner.TextStateCharacterSpacingTargetForListLabel(decimalLabel, 12d);
        DocxTextStateCharacterSpacingTarget bulletTarget = DocxTextEmissionPlanner.TextStateCharacterSpacingTargetForListLabel(bulletLabel, 12d);
        DocxTextEmissionPlan decimalPlan = DocxTextEmissionPlanner.CreateForListLabel(decimalRun, decimalLabel);
        DocxTextEmissionPlan bulletPlan = DocxTextEmissionPlanner.CreateForListLabel(bulletRun, bulletLabel);

        TestAssert.Equal(0.048d, DocxTextEmissionPlanner.TextStateCharacterSpacingForListLabel(decimalLabel, 12d));
        TestAssert.Equal(0d, DocxTextEmissionPlanner.TextStateCharacterSpacingForListLabel(bulletLabel, 12d));
        TestAssert.Equal(0.048d, decimalTarget.CharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, decimalTarget.Source);
        TestAssert.Equal(0d, bulletTarget.CharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, bulletTarget.Source);
        TestAssert.Equal(0.048d, decimalPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, decimalPlan.PdfCharacterSpacingSource);
        TestAssert.True(!decimalPlan.CompensatePdfCharacterSpacing, "List-label PDF text state should be emitted, not folded into layout positioning.");
        TestAssert.Equal(0d, bulletPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.ListLabel, bulletPlan.PdfCharacterSpacingSource);
    }

    public static void DocxTextEmissionPlannerOwnsPdfTextStateAndPositioningSpacing()
    {
        var run = new DocxTextRun("Tracked", 11d, null, false, false, false, null, null, CharacterSpacingPoints: 0.25d);

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.Create(run, 11d, pdfCharacterSpacing: 0.05d, compensatePdfCharacterSpacing: true);

        TestAssert.Equal(11.04d, plan.PdfFontSize);
        TestAssert.Equal(0.05d, plan.PdfCharacterSpacing);
        TestAssert.Equal(0.20d, plan.PositioningCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.Explicit, plan.PdfCharacterSpacingSource);
        TestAssert.True(plan.CompensatePdfCharacterSpacing, "Planner should record that positioned glyph advances compensate PDF Tc.");
    }

    public static void DocxTextEmissionPlannerKeepsNumberedLabelTcOutOfPositioning()
    {
        var run = new DocxTextRun("1", 12d, null, false, false, false, null, null);
        double numberedTc = OfficePdfTextEmissionProfile.ObservedWordNumberedListTextStateCharacterSpacing(12d);

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.Create(run, 12d, numberedTc, compensatePdfCharacterSpacing: false);

        TestAssert.Equal(12d, plan.PdfFontSize);
        TestAssert.Equal(numberedTc, plan.PdfCharacterSpacing);
        TestAssert.Equal(0d, plan.PositioningCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.Explicit, plan.PdfCharacterSpacingSource);
        TestAssert.True(!plan.CompensatePdfCharacterSpacing, "Numbered labels use PDF Tc as emitted text state, not as a compensated layout offset.");
    }

    public static void DocxTextEmissionPlannerForcesTerminalLineSpacesToNeutralTc()
    {
        var run = new DocxTextRun("Body", 11d, null, false, false, false, null, null, CharacterSpacingPoints: 0.25d);

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateTerminalLineSpace(run, 11d);

        TestAssert.Equal(11.04d, plan.PdfFontSize);
        TestAssert.Equal(0d, plan.PdfCharacterSpacing);
        TestAssert.Equal(0.25d, plan.PositioningCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.TerminalLineSpace, plan.PdfCharacterSpacingSource);
        TestAssert.True(plan.CompensatePdfCharacterSpacing, "Terminal spaces should stay eligible for authored positioning while emitting neutral PDF Tc.");
    }

    public static void DocxTextEmissionPlannerOwnsTerminalSegmentPlanOverride()
    {
        var run = new DocxTextRun("Body", 11d, null, false, false, false, null, null, CharacterSpacingPoints: 0.25d);

        DocxTextEmissionPlan textPlan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            run,
            11d,
            pdfCharacterSpacing: 0.05d,
            DocxTextStateCharacterSpacingSource.AdvanceTarget,
            compensatePdfCharacterSpacing: false,
            isTerminalLineSpace: false);
        DocxTextEmissionPlan terminalPlan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            run,
            11d,
            pdfCharacterSpacing: 0.05d,
            DocxTextStateCharacterSpacingSource.AdvanceTarget,
            compensatePdfCharacterSpacing: false,
            isTerminalLineSpace: true);

        TestAssert.Equal(0.05d, textPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, textPlan.PdfCharacterSpacingSource);
        TestAssert.True(!textPlan.CompensatePdfCharacterSpacing, "Non-terminal segments should preserve the caller's compensation mode.");
        TestAssert.Equal(0d, terminalPlan.PdfCharacterSpacing);
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.TerminalLineSpace, terminalPlan.PdfCharacterSpacingSource);
        TestAssert.True(terminalPlan.CompensatePdfCharacterSpacing, "Terminal spaces should always be planned through the neutral terminal-space path.");
        TestAssert.Equal(0.25d, terminalPlan.PositioningCharacterSpacing);
    }

    public static void DocxTextEmissionPlannerDerivesTcFromAdvanceTarget()
    {
        var run = new DocxTextRun("Body", 11d, null, false, false, false, null, null, CharacterSpacingPoints: 0.12d);

        double tc = DocxTextEmissionPlanner.TextStateCharacterSpacingForAdvanceTarget(
            glyphGapCount: 3,
            currentEmittedAdvance: 24d,
            targetEmittedAdvance: 24.18d);
        DocxTextEmissionPlan emittedPlan = DocxTextEmissionPlanner.CreateForAdvanceTarget(
            run,
            11d,
            glyphGapCount: 3,
            currentEmittedAdvance: 24d,
            targetEmittedAdvance: 24.18d,
            compensatePdfCharacterSpacing: false);
        DocxTextEmissionPlan compensatedPlan = DocxTextEmissionPlanner.CreateForAdvanceTarget(
            run,
            11d,
            glyphGapCount: 3,
            currentEmittedAdvance: 24d,
            targetEmittedAdvance: 24.18d,
            compensatePdfCharacterSpacing: true);

        TestAssert.True(Math.Abs(tc - 0.06d) < 0.0001d, "Tc should be the emitted-advance delta distributed over glyph gaps.");
        TestAssert.True(Math.Abs(emittedPlan.PdfCharacterSpacing - 0.06d) < 0.0001d, "Uncompensated plans should carry the derived Tc.");
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, emittedPlan.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(emittedPlan.PositioningCharacterSpacing - 0.12d) < 0.0001d, "Uncompensated plans should leave positioning spacing unchanged.");
        TestAssert.True(Math.Abs(compensatedPlan.PdfCharacterSpacing - 0.06d) < 0.0001d, "Compensated plans should carry the same derived Tc.");
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, compensatedPlan.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(compensatedPlan.PositioningCharacterSpacing - 0.06d) < 0.0001d, "Compensated plans should subtract derived Tc from positioning spacing.");
        TestAssert.Equal(0d, DocxTextEmissionPlanner.TextStateCharacterSpacingForAdvanceTarget(0, 24d, 24.18d));
    }

    public static void DocxTextEmissionPlannerSplitsDashPunctuationIntoOperationParts()
    {
        var run = new DocxTextRun("Alpha-Beta", 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout("Alpha-Beta", run, 20d, 100d);

        IReadOnlyList<DocxTextEmissionPart> parts = DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, 10d, new FontSizeWidthTextMeasurer());

        TestAssert.Equal(3, parts.Count);
        TestAssert.Equal("Alpha", parts[0].Text);
        TestAssert.Equal("-", parts[1].Text);
        TestAssert.Equal("Beta", parts[2].Text);
        TestAssert.Equal(20d, parts[0].X);
        TestAssert.Equal(70d, parts[1].X);
        TestAssert.Equal(80d, parts[2].X);
        TestAssert.Equal(40d, parts[2].Width);
    }

    public static void DocxTextEmissionPlannerKeepsWholeOperationWithoutMeasurer()
    {
        var run = new DocxTextRun("Alpha-Beta", 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout("Alpha-Beta", run, 20d, 50d);

        IReadOnlyList<DocxTextEmissionPart> parts = DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, 10d, null);

        TestAssert.Equal(1, parts.Count);
        TestAssert.Equal("Alpha-Beta", parts[0].Text);
        TestAssert.Equal(20d, parts[0].X);
        TestAssert.Equal(50d, parts[0].Width);
    }

    public static void DocxTextEmissionPlannerClassifiesTextWithoutExposingIt()
    {
        DocxTextEmissionCharacterProfile profile = DocxTextEmissionPlanner.ClassifyText("A9 -+");

        TestAssert.Equal(1, profile.LetterCount);
        TestAssert.Equal(1, profile.DigitCount);
        TestAssert.Equal(1, profile.WhitespaceCount);
        TestAssert.Equal(1, profile.PunctuationCount);
        TestAssert.Equal(1, profile.SymbolCount);
        TestAssert.Equal(0, profile.OtherCount);
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
        TestAssert.Equal(DocxTextAlignment.Justified, document.Paragraphs[1].Alignment);
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
                      <w:ind w:left="720" w:start="960" w:right="240" w:end="480" w:hanging="360"/>
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
                        <w:ind w:firstLine="120"/>
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
        TestAssert.Equal(14d, paragraph.SpacingAfterPoints);
        TestAssert.Equal("360", paragraph.Spacing.BeforeValue ?? string.Empty);
        TestAssert.Equal(string.Empty, paragraph.Spacing.AfterValue ?? string.Empty);
        TestAssert.Equal(string.Empty, paragraph.Spacing.AfterLinesValue ?? string.Empty);
        TestAssert.Equal(string.Empty, paragraph.Spacing.BeforeAutoSpacingValue ?? string.Empty);
        TestAssert.Equal("1", paragraph.Spacing.AfterAutoSpacingValue ?? string.Empty);
        TestAssert.Equal("480", paragraph.Spacing.LineValue ?? string.Empty);
        TestAssert.Equal("exact", paragraph.Spacing.LineRuleValue ?? string.Empty);
        TestAssert.Equal(48d, paragraph.Indent.LeftPoints ?? 0d);
        TestAssert.Equal(24d, paragraph.Indent.RightPoints ?? 0d);
        TestAssert.Equal(6d, paragraph.Indent.FirstLinePoints ?? 0d);
        TestAssert.Equal("960", paragraph.Indent.LeftValue ?? string.Empty);
        TestAssert.Equal("480", paragraph.Indent.RightValue ?? string.Empty);
        TestAssert.Equal("120", paragraph.Indent.FirstLineValue ?? string.Empty);
        TestAssert.Equal(string.Empty, paragraph.Indent.HangingValue ?? string.Empty);
        TestAssert.True(paragraph.Spacing.ContextualSpacing == true, "Style contextual spacing should survive the paragraph cascade.");
        TestAssert.True(paragraph.KeepRules.KeepNext == true, "Style keepNext should survive the paragraph cascade.");
        TestAssert.True(paragraph.KeepRules.KeepLines == false, "Explicit off keepLines should survive the paragraph cascade.");
        TestAssert.True(paragraph.KeepRules.WidowControl == true, "Style widowControl should survive the paragraph cascade.");
    }

    public static void DocxReaderCascadesParagraphSpacingAndRunSizeThroughBasedOnStyles()
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
                    <w:pPr><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/></w:pPr>
                    <w:rPr><w:sz w:val="28"/><w:color w:val="112233"/></w:rPr>
                  </w:style>
                  <w:style w:type="paragraph" w:styleId="Child">
                    <w:basedOn w:val="Base"/>
                    <w:pPr><w:spacing w:after="0"/></w:pPr>
                    <w:rPr><w:b/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr><w:pStyle w:val="Child"/></w:pPr>
                      <w:r><w:t>Inherited spacing and size</w:t></w:r>
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
        DocxTextRun run = paragraph.Runs.Single();
        TestAssert.Equal("Child", paragraph.StyleId ?? string.Empty);
        TestAssert.Equal(6d, paragraph.SpacingBeforePoints);
        TestAssert.Equal(0d, paragraph.SpacingAfterPoints);
        TestAssert.Equal("120", paragraph.Spacing.BeforeValue ?? string.Empty);
        TestAssert.Equal("0", paragraph.Spacing.AfterValue ?? string.Empty);
        TestAssert.Equal("360", paragraph.Spacing.LineValue ?? string.Empty);
        TestAssert.Equal("auto", paragraph.Spacing.LineRuleValue ?? string.Empty);
        TestAssert.Equal(1.5d, paragraph.LineSpacingFactor);
        TestAssert.Equal(14d, run.FontSize);
        TestAssert.Equal("112233", run.ColorHex ?? string.Empty);
        TestAssert.True(run.Bold, "Child paragraph style run properties should merge over the inherited base run size/color.");
    }

    public static void DocxReaderPreservesEmptyParagraphBodyElement()
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
                    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
                    <w:p/>
                    <w:p><w:r><w:t>Beta</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxParagraphElement[] paragraphs = document.BodyElements.OfType<DocxParagraphElement>().ToArray();
        TestAssert.Equal(3, paragraphs.Length);
        TestAssert.Equal(1, paragraphs[1].Paragraph.Runs.Count);
        TestAssert.Equal(string.Empty, paragraphs[1].Paragraph.Runs[0].Text);
    }

    public static void DocxReaderPreservesDocumentSettingsCompatibilityFacts()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
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
                  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
                </Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:defaultTabStop w:val="720"/>
                  <w:characterSpacingControl w:val="doNotCompress"/>
                  <w:compat>
                    <w:useFELayout/>
                    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
                  </w:compat>
                </w:settings>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Settings</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.Equal("doNotCompress", document.Settings.CharacterSpacingControlValue ?? string.Empty);
        TestAssert.Equal("720", document.Settings.DefaultTabStopValue ?? string.Empty);
        TestAssert.Equal(36d, document.Settings.DefaultTabStopPoints ?? 0d);
        TestAssert.True(document.Settings.UseFELayout == true, "Empty useFELayout should opt in.");
        DocxCompatSetting compat = document.Settings.CompatSettings.Single();
        TestAssert.Equal("compatibilityMode", compat.Name ?? string.Empty);
        TestAssert.Equal("15", compat.Value ?? string.Empty);
    }

    public static void DocxLayoutStageEmitsEmptyParagraphMarkSpaceLine()
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
                    <w:p><w:pPr><w:spacing w:after="0" w:line="240"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>
                    <w:p><w:pPr><w:spacing w:after="0" w:line="240"/></w:pPr></w:p>
                    <w:p><w:pPr><w:spacing w:after="0" w:line="240"/></w:pPr><w:r><w:t>Beta</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(3, lines.Length);
        TestAssert.Equal("Alpha", lines[0].Text);
        TestAssert.Equal(" ", lines[1].Text);
        TestAssert.Equal("Beta", lines[2].Text);
        TestAssert.Equal(22d, Math.Round(lines[0].BaselineY - lines[2].BaselineY, 3));
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
        TestAssert.Equal(36d, paragraph.SpacingBeforePoints);
        TestAssert.Equal(48d, paragraph.SpacingAfterPoints);
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

    public static void DocxReaderUsesWordDefaultAutoLineSpacingWhenLineTokenIsAbsent()
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
                      <w:pPr><w:spacing w:before="36" w:after="0"/></w:pPr>
                      <w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Missing line token</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>No spacing token</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.Equal(2, document.Paragraphs.Count);
        TestAssert.Equal(1.2d, document.Paragraphs[0].LineSpacingFactor);
        TestAssert.Equal(1.2d, document.Paragraphs[1].LineSpacingFactor);
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
                    <w:charset w:val="00"/>
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
        TestAssert.Equal("00", entry.CharsetValue ?? string.Empty);
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
            [new DocxFontTableEntry("Corporate Sans", "Aptos", "swiss", "variable", null, null)],
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
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null, null)],
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
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null, null)],
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

    public static void DocxFontPlanUsesBodyElementInventoryAsCanonicalSource()
    {
        var bodyRun = new DocxTextRun("Body", 11d, null, false, false, false, null, "Body Sans")
        {
            Fonts = new DocxRunFonts("Body Sans", null, null, null, null, null, null, null)
        };
        DocxParagraph bodyParagraph = CreateFontPlanParagraph(bodyRun);
        var cellRun = new DocxTextRun("Cell", 11d, null, false, false, false, null, "Cell Sans")
        {
            Fonts = new DocxRunFonts("Cell Sans", null, null, null, null, null, null, null)
        };
        DocxParagraph cellParagraph = CreateFontPlanParagraph(cellRun);
        var table = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [cellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], null)]);
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
            [new DocxParagraphElement(bodyParagraph), new DocxTableElement(table)],
            [],
            []);
        var resolver = new MapFontResolver(["Body Sans", "Cell Sans"], "Resolver Fallback");

        string plannedTexts = string.Join("|", DocxFontPlan.Create(document, resolver).Runs.Select(run => run.Run.Text).Order(StringComparer.Ordinal));

        TestAssert.Equal("Body|Cell", plannedTexts);
    }

    public static void DocxBlockTraversalAndFontPlanIncludeNestedTableCellBody()
    {
        DocxParagraph nestedParagraph = CreateDocxLayoutParagraph("Nested cell", 12d, 12d);
        var nestedTable = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxTableCell outerCell = new(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var outerTable = new DocxTable(null, [60d], [new DocxTableRow([outerCell], 30d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(outerTable)], [outerTable]);
        var resolver = new MapFontResolver(["Arial"], "Resolver Fallback");

        DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(document).ToArray();
        DocxTable[] tables = DocxBlockTraversal.EnumerateBodyTables(document).ToArray();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver);

        TestAssert.Equal(1, paragraphs.Length);
        TestAssert.Equal("Nested cell", paragraphs[0].Runs.Single().Text);
        TestAssert.Equal(2, tables.Length);
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Nested cell"), "Nested table-cell body text should participate in DOCX font planning.");
    }

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
        DocxDocument document = CreateFontPlanDocument([paragraph], DocxFontCatalog.Empty);
        var resolver = new MapFontResolver(["Body Sans", "Marker Sans"], "Resolver Fallback");

        DocxResolvedRunTypeface marker = DocxFontPlan.Create(document, resolver).Runs.Single(run => run.Run.Text == "#");

        TestAssert.Equal("Marker Sans", marker.RequestedFamily ?? string.Empty);
        TestAssert.Equal(DocxTypefaceResolutionSource.Primary, marker.Source);
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
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null, null)],
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

    public static void DocxFontPlanSnapshotReportsPrivateSafeOpenTypeMetrics()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        var run = new DocxTextRun("Metric probe", 10d, null, false, false, false, null, font.Value.Resolution.FamilyName)
        {
            Fonts = new DocxRunFonts(font.Value.Resolution.FamilyName, null, null, null, null, null, null, null)
        };
        DocxDocument document = CreateFontPlanDocument(run, new DocxFontCatalog([], DocxThemeFonts.Empty));

        DocxFontPlanSnapshot snapshot = new DocxRenderer(new SingleResolutionFontResolver(font.Value.Resolution)).InspectFontPlan(document);

        DocxFontMetricBucketSnapshot bucket = snapshot.MetricBuckets.Single();
        TestAssert.Equal(DocxTypefaceResolutionSource.Primary.ToString(), bucket.Source);
        TestAssert.Equal(10d, bucket.FontSize);
        TestAssert.Equal(1, bucket.RunCount);
        TestAssert.Equal(12, bucket.ResolvedFamilyHash?.Length ?? 0);
        TestAssert.Equal(font.Value.Font.UnitsPerEm, bucket.UnitsPerEm ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.TypographicAscender, bucket.TypographicAscender ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.TypographicDescender, bucket.TypographicDescender ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.TypographicLineGap, bucket.TypographicLineGap ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.WindowsAscender, bucket.WindowsAscender ?? 0);
        TestAssert.Equal(font.Value.Font.Os2.WindowsDescender, bucket.WindowsDescender ?? 0);
        TestAssert.Equal(DocxLineMetrics.MeasureOpenTypeSingleLineHeight(font.Value.Font, 10d), bucket.SingleLineHeightPoints ?? 0d);
        TestAssert.Equal(DocxLineMetrics.MeasureWindowsAscender(font.Value.Font, 10d), bucket.WindowsAscenderPoints ?? 0d);
        TestAssert.Equal(DocxLineMetrics.MeasureWindowsDescender(font.Value.Font, 10d), bucket.WindowsDescenderPoints ?? 0d);
    }

    public static void DocxStructureSnapshotReportsPreLayoutBlockAndTableFacts()
    {
        DocxParagraph paragraph = CreateDocxLayoutParagraph("A1 body.", fontSize: 10d, lineSpacingPoints: 12d) with
        {
            StyleId = "BodyStyle",
            Indent = new DocxParagraphIndent(12d, 3d, null, 6d, "240", "60", null, "120"),
            TabStops = [new DocxTabStop(36d, "720", "left", null)],
            SnapToGrid = true,
            SnapToGridValue = "1"
        };
        DocxParagraph breakParagraph = CreateDocxLayoutParagraph("Break", fontSize: 9d, lineSpacingPoints: 11d);
        var listLabel = new DocxListLabel(
            "1",
            "decimal",
            "%1.",
            "tab",
            "7",
            0,
            DocxNumberingIndent.Empty,
            DocxTextRunStyle.Empty);
        DocxParagraph cellParagraph = CreateDocxLayoutParagraph("Cell 42", fontSize: 9d, lineSpacingPoints: 10d) with
        {
            StyleId = "CellStyle",
            ListLabel = listLabel
        };
        var restartCell = new DocxTableCell(
            string.Empty,
            [cellParagraph],
            "D9EAF7",
            "clear",
            null,
            "center",
            [new DocxTableCellBorder("bottom", "single", "auto", "4")],
            DocxTableCellMargins.Empty,
            GridSpan: 2,
            GridSpanValue: "2",
            HasVerticalMerge: true,
            VerticalMergeValue: "restart");
        var continuationCell = new DocxTableCell(
            string.Empty,
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty,
            HasVerticalMerge: true,
            VerticalMergeValue: "continue");
        var table = new DocxTable(
            "fixed",
            [20d, 20d],
            [
                new DocxTableRow([restartCell], 18d, IsHeader: true, HeaderValue: "1", HeightValue: "360", HeightRuleValue: "atLeast", CantSplit: true, CantSplitValue: "1"),
                new DocxTableRow([continuationCell], null)
            ],
            StyleId: "TableGrid",
            PreferredWidthPoints: 40d,
            PreferredWidthValue: "800",
            PreferredWidthType: "dxa",
            IndentPoints: 6d,
            IndentValue: "120",
            IndentType: "dxa",
            CellSpacingPoints: 1d,
            CellSpacingValue: "20",
            CellSpacingType: "dxa",
            Look: new DocxTableLook(null, true, "1", null, null, true, "1", null, null, null, null, null, null));
        DocxParagraph documentHeader = CreateDocxLayoutParagraph("Header", fontSize: 9d, lineSpacingPoints: 10d);
        DocxParagraph documentFooter = CreateDocxLayoutParagraph("Footer", fontSize: 9d, lineSpacingPoints: 10d);
        DocxParagraph sectionHeader = CreateDocxLayoutParagraph("Section", fontSize: 9d, lineSpacingPoints: 10d);
        DocxPageSettings sectionSettings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["first"] = [sectionHeader]
            }
        };
        var sectionBreak = new DocxSectionBreakElement(sectionSettings, "nextPage", "2", "1", "720", []);
        var floatingDrawing = new DocxFloatingDrawing(
            "0",
            "0",
            "114300",
            "114300",
            "0",
            "251659264",
            "0",
            "0",
            "1",
            "1",
            "914400",
            "457200",
            "column",
            "center",
            null,
            "paragraph",
            null,
            "12700",
            "square",
            "bothSides");
        var document = new DocxDocument(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [
                new DocxParagraphElement(paragraph),
                new DocxPageBreakElement("runBreak", "page", breakParagraph),
                new DocxTableElement(table),
                sectionBreak
            ],
            [paragraph],
            [table])
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [documentHeader]
            },
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["even"] = [documentFooter]
            }
        };

        DocxStructureSnapshot snapshot = new DocxRenderer(new MapFontResolver([], "Fallback")).InspectStructure(document);

        TestAssert.Equal(4, snapshot.BlockCount);
        TestAssert.Equal(1, snapshot.ParagraphBlockCount);
        TestAssert.Equal(1, snapshot.PageBreakBlockCount);
        TestAssert.Equal(0, snapshot.ManualBreakBlockCount);
        TestAssert.Equal(1, snapshot.TableBlockCount);
        TestAssert.Equal(1, snapshot.SectionBreakBlockCount);
        TestAssert.Equal(8, snapshot.BodyTextLength);
        TestAssert.Equal(1, snapshot.FloatingDrawingCount);
        TestAssert.Equal("Paragraph", snapshot.Blocks[1].PreviousKind ?? string.Empty);
        TestAssert.Equal("Table", snapshot.Blocks[1].NextKind ?? string.Empty);
        TestAssert.True(snapshot.Blocks[0].SnapToGrid == true, "Paragraph structure should expose snapToGrid before layout.");
        TestAssert.Equal(1, snapshot.Blocks[0].TabStopCount ?? 0);
        TestAssert.Equal(2, snapshot.Blocks[0].WhitespaceDelimitedTokenCount ?? 0);
        TestAssert.Equal(5, snapshot.Blocks[0].LongestWhitespaceDelimitedTokenLength ?? 0);
        TestAssert.True(snapshot.Blocks[1].PageBreakConsumesParagraphLine == true, "Page-break structure should expose the consumed paragraph line.");
        TestAssert.Equal("nextPage", snapshot.Blocks[3].SectionBreakTypeValue ?? string.Empty);
        TestAssert.Equal("2", snapshot.Blocks[3].SectionColumnCountValue ?? string.Empty);
        TestAssert.Equal("1", snapshot.Blocks[3].SectionColumnEqualWidthValue ?? string.Empty);
        TestAssert.Equal("720", snapshot.Blocks[3].SectionColumnSpaceValue ?? string.Empty);
        TestAssert.Equal(4, snapshot.Stories.Count);
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Body" && story.BlockCount == 4), "Body story should summarize the document block stream.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Header" && story.Scope == "document" && story.VariantType == "default"), "Document default header story should be inventoried before layout.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Footer" && story.Scope == "document" && story.VariantType == "even"), "Document even footer story should be inventoried before layout.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Header" && story.Scope == "section@3" && story.SectionBreakBlockIndex == 3 && story.VariantType == "first"), "Section header story should be tied to its section-break block.");
        TestAssert.True(snapshot.StyleUsages.Any(usage => usage.Kind == "Paragraph" && usage.StyleId == "BodyStyle" && usage.ParagraphCount == 1), "Paragraph style usage should be available before layout.");
        TestAssert.True(snapshot.StyleUsages.Any(usage => usage.Kind == "Paragraph" && usage.StyleId == "CellStyle" && usage.ParagraphCount == 1), "Table-cell paragraph style usage should be available before layout.");
        TestAssert.True(snapshot.StyleUsages.Any(usage => usage.Kind == "Table" && usage.StyleId == "TableGrid" && usage.TableCount == 1), "Table style usage should be available before layout.");
        DocxStructureListUsageSnapshot listUsage = snapshot.ListUsages.Single();
        TestAssert.Equal("7", listUsage.NumberId);
        TestAssert.Equal("decimal", listUsage.FormatValue);
        TestAssert.Equal(1, listUsage.ParagraphCount);
        DocxStructureFloatingDrawingSnapshot drawingSnapshot = snapshot.FloatingDrawings.Single();
        TestAssert.Equal("square", drawingSnapshot.WrapKind ?? string.Empty);
        TestAssert.Equal("column", drawingSnapshot.HorizontalRelativeFromValue ?? string.Empty);
        TestAssert.Equal("paragraph", drawingSnapshot.VerticalRelativeFromValue ?? string.Empty);

        DocxStructureTableSnapshot tableSnapshot = snapshot.Tables.Single();
        TestAssert.Equal(2, tableSnapshot.RowCount);
        TestAssert.Equal(2, tableSnapshot.MaxColumnCount);
        TestAssert.Equal(1, tableSnapshot.HeaderRowCount);
        TestAssert.Equal(1, tableSnapshot.CantSplitRowCount);
        TestAssert.Equal(1, tableSnapshot.DeclaredHeightRowCount);
        TestAssert.Equal(1, tableSnapshot.AtLeastHeightRowCount);
        TestAssert.Equal(1, tableSnapshot.GridSpanCellCount);
        TestAssert.Equal(2, tableSnapshot.VerticalMergeCellCount);
        TestAssert.Equal(1, tableSnapshot.VerticalMergeRestartCellCount);
        TestAssert.Equal(1, tableSnapshot.ShadedCellCount);
        TestAssert.Equal(1, tableSnapshot.VisibleBorderCount);
        TestAssert.Equal(1, tableSnapshot.ParagraphCount);
        TestAssert.Equal(7, tableSnapshot.TextLength);
        TestAssert.Equal(2, tableSnapshot.WhitespaceDelimitedTokenCount);
        TestAssert.Equal(4, tableSnapshot.LongestWhitespaceDelimitedTokenLength);
        TestAssert.True(tableSnapshot.LookFirstRow == true, "Table look facts should be present before rendering.");
        TestAssert.Equal("PageBreak", snapshot.TableAdjacency.Single().PreviousKind ?? string.Empty);
        DocxStructureTableRowSnapshot rowSnapshot = tableSnapshot.Rows[0];
        TestAssert.True(rowSnapshot.IsHeader, "Row profile should expose header rows before pagination.");
        TestAssert.True(rowSnapshot.CantSplit, "Row profile should expose cantSplit before pagination.");
        TestAssert.Equal(1, rowSnapshot.GridSpanCellCount);
        TestAssert.Equal(1, rowSnapshot.VerticalMergeRestartCellCount);
        TestAssert.Equal(1, rowSnapshot.ShadedCellCount);
        TestAssert.Equal(1, rowSnapshot.VisibleBorderCount);
        TestAssert.Equal(7, rowSnapshot.TextLength);
        TestAssert.Equal(2, rowSnapshot.WhitespaceDelimitedTokenCount);
        TestAssert.Equal(4, rowSnapshot.LongestWhitespaceDelimitedTokenLength);
        DocxStructureTableCellSnapshot cellSnapshot = rowSnapshot.Cells[0];
        TestAssert.Equal(2, cellSnapshot.GridSpan);
        TestAssert.True(cellSnapshot.HasVerticalMerge, "Cell profile should expose vertical merge state before layout.");
        TestAssert.Equal(2, cellSnapshot.DigitCharacterCount);
        TestAssert.Equal(2, cellSnapshot.WhitespaceDelimitedTokenCount);
        TestAssert.Equal(4, cellSnapshot.LongestWhitespaceDelimitedTokenLength);
    }

    public static void DocxStructureSnapshotNormalizesPlainTableCellText()
    {
        DocxTable table = CreateSingleCellTable("Plain 123", 20d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxStructureSnapshot snapshot = new DocxRenderer().InspectStructure(document);

        DocxStructureTableSnapshot tableSnapshot = snapshot.Tables.Single();
        DocxStructureTableCellSnapshot cellSnapshot = tableSnapshot.Rows.Single().Cells.Single();
        TestAssert.Equal(1, tableSnapshot.ParagraphCount);
        TestAssert.Equal(9, tableSnapshot.TextLength);
        TestAssert.Equal(2, tableSnapshot.WhitespaceDelimitedTokenCount);
        TestAssert.Equal(5, tableSnapshot.LongestWhitespaceDelimitedTokenLength);
        TestAssert.Equal(1, cellSnapshot.ParagraphCount);
        TestAssert.Equal(1, cellSnapshot.RunCount);
        TestAssert.Equal(9, cellSnapshot.TextLength);
        TestAssert.Equal(3, cellSnapshot.DigitCharacterCount);
        TestAssert.Equal(1, cellSnapshot.SpaceCharacterCount);
        TestAssert.True(snapshot.StyleUsages.Any(usage => usage.Kind == "Paragraph" && usage.StyleId is null && usage.ParagraphCount == 1 && usage.TextLength == 9), "Plain table-cell text should contribute to paragraph style usage through the shared cell content stream.");
    }

    public static void DocxStructureSnapshotReportsTableCellHyperlinkInventory()
    {
        DocxParagraph cellParagraph = CreateDocxLayoutParagraph("External Internal", fontSize: 10d, lineSpacingPoints: 12d) with
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdExt", null, null, null, "https://example.invalid/", "External", null, 0, 1, 0, 1, 8),
                new DocxHyperlinkSpan(null, "Bookmark", null, null, null, null, null, 1, 1, 1, 1, 8)
            ]
        };
        var table = new DocxTable(
            null,
            [80d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [cellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxStructureTableSnapshot tableSnapshot = new DocxRenderer()
            .InspectStructure(document)
            .Tables
            .Single();
        DocxStructureTableRowSnapshot rowSnapshot = tableSnapshot.Rows.Single();
        DocxStructureTableCellSnapshot cellSnapshot = rowSnapshot.Cells.Single();

        TestAssert.Equal(2, tableSnapshot.HyperlinkCount);
        TestAssert.Equal(1, tableSnapshot.ExternalHyperlinkCount);
        TestAssert.Equal(1, tableSnapshot.InternalHyperlinkCount);
        TestAssert.Equal(2, rowSnapshot.HyperlinkCount);
        TestAssert.Equal(1, rowSnapshot.ExternalHyperlinkCount);
        TestAssert.Equal(1, rowSnapshot.InternalHyperlinkCount);
        TestAssert.Equal(2, cellSnapshot.HyperlinkCount);
        TestAssert.Equal(1, cellSnapshot.ExternalHyperlinkCount);
        TestAssert.Equal(1, cellSnapshot.InternalHyperlinkCount);
    }

    public static void DocxRendererEmitsBodyExternalHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Before ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null),
            new DocxTextRun(" After", 10d, null, false, false, false, null, null)
        };
        var paragraph = new DocxParagraph(
            runs,
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdLink", null, null, null, "https://example.invalid/docx", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        DocxDocument document = CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/docx", annotation.Uri);
        TestAssert.True(annotation.X > document.MarginLeftPoints, "The annotation should be anchored to the placed hyperlink run, not the paragraph origin.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover the rendered hyperlink text.");
        TestAssert.True(annotation.Height > 0d, "The annotation should use font metrics for a non-empty rectangle.");
    }

    public static void DocxRendererDoesNotEmitUriAnnotationsForInternalHyperlinks()
    {
        DocxParagraph paragraph = CreateDocxLayoutParagraph("Internal", 10d, 12d) with
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "Bookmark", null, null, null, null, null, 0, 1, 0, 1, 8)
            ]
        };
        DocxDocument document = CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        TestAssert.Equal(0, page.Annotations.Count);
    }

    public static void DocxRendererEmitsInternalHyperlinkDestinationsFromBookmarks()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Target", 10d, null, false, false, false, null, null)
            ],
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "BookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 6)
            ]
        };
        DocxParagraph targetParagraph = CreateDocxLayoutParagraph("Target text", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("3", "BookmarkTarget", 0, 0, 0)
            ]
        };
        DocxDocument document = CreateLayoutTestDocument(
            [new DocxParagraphElement(linkParagraph), new DocxParagraphElement(targetParagraph)],
            []);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal DOCX links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal DOCX links should resolve to a PDF page destination.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed bookmark text coordinates.");
        TestAssert.True(annotation.Destination?.Top > 0d, "The destination should point to a concrete bookmark line top.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should still cover the rendered internal-link text.");
    }

    public static void DocxRendererEmitsTableCellExternalHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Cell ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var paragraph = new DocxParagraph(
            runs,
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdCell", null, null, null, "https://example.invalid/cell", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        var table = new DocxTable(
            null,
            [100d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 24d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/cell", annotation.Uri);
        TestAssert.True(annotation.X > document.MarginLeftPoints, "The annotation should be anchored to the placed table-cell hyperlink run.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover table-cell hyperlink text.");
    }

    public static void DocxRendererEmitsTableCellInternalHyperlinkDestinations()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Cell ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Target", 10d, null, false, false, false, null, null)
            ],
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "CellBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 6)
            ]
        };
        DocxParagraph targetParagraph = CreateDocxLayoutParagraph("Bookmark target", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("7", "CellBookmarkTarget", 0, 0, 0)
            ]
        };
        var table = new DocxTable(
            null,
            [120d],
            [
                new DocxTableRow(
                    [new DocxTableCell(string.Empty, [linkParagraph, targetParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)],
                    36d)
            ]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal table-cell links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal table-cell links should resolve to a PDF page destination.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed table-cell bookmark coordinates.");
        TestAssert.True(annotation.Destination?.Top > 0d, "The destination should point to a concrete table-cell bookmark line top.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should cover the rendered table-cell internal-link text.");
    }

    public static void DocxHeaderRendererEmitsExternalHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Header ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var header = new DocxParagraph(
            runs,
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdHeader", null, null, null, "https://example.invalid/header", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            }
        };
        DocxParagraph body = CreateDocxLayoutParagraph("Body", 10d, 12d);
        var document = new DocxDocument(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            pageSettings,
            [],
            [header],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/header", annotation.Uri);
        TestAssert.True(annotation.Y > 150d, "Header annotation should be anchored near the page top static story.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover header hyperlink text.");
    }

    public static void DocxFooterRendererEmitsExternalHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Footer ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var footer = new DocxParagraph(
            runs,
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdFooter", null, null, null, "https://example.invalid/footer", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [footer]
            }
        };
        DocxParagraph body = CreateDocxLayoutParagraph("Body", 10d, 12d);
        var document = new DocxDocument(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            pageSettings,
            [],
            [],
            [footer],
            [new DocxParagraphElement(body)],
            [body],
            []);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/footer", annotation.Uri);
        TestAssert.True(annotation.Y < 60d, "Footer annotation should be anchored near the page bottom static story.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover footer hyperlink text.");
    }

    public static void DocxStaticStoryRendererEmitsInternalHyperlinkDestinations()
    {
        DocxParagraph header = new(
            [
                new DocxTextRun("Header ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Jump", 10d, null, false, false, false, null, null)
            ],
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "StaticBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 4)
            ]
        };
        DocxParagraph target = CreateDocxLayoutParagraph("Body target", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("12", "StaticBookmarkTarget", 0, 0, 0)
            ]
        };
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            }
        };
        var document = new DocxDocument(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            pageSettings,
            [],
            [header],
            [],
            [new DocxParagraphElement(target)],
            [target],
            []);

        PdfPage page = new DocxRenderer().RenderBlankPages(document).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal static-story links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal static-story links should resolve through the shared bookmark destination map.");
        TestAssert.True(annotation.Y > 150d, "The clickable rectangle should be anchored to the header story text.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed body bookmark coordinates.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover static-story hyperlink text.");
    }

    public static void DocxStructureSnapshotUsesBodyElementInventoryAsCanonicalSource()
    {
        DocxParagraph cellParagraph = CreateDocxLayoutParagraph("Cell", fontSize: 9d, lineSpacingPoints: 10d) with
        {
            StyleId = "CellStyle"
        };
        var table = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [cellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], null)],
            StyleId: "TableStyle");
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
            [new DocxTableElement(table)],
            [],
            []);

        DocxStructureSnapshot snapshot = DocxStructureSnapshot.FromDocument(document);

        TestAssert.Equal(1, snapshot.TableBlockCount);
        TestAssert.True(snapshot.StyleUsages.Any(usage => usage.Kind == "Table" && usage.StyleId == "TableStyle" && usage.TableCount == 1), "Table style usage should come from the body block stream.");
        TestAssert.True(snapshot.StyleUsages.Any(usage => usage.Kind == "Paragraph" && usage.StyleId == "CellStyle" && usage.ParagraphCount == 1), "Cell paragraph style usage should come from the body block stream.");
    }

    public static void DocxFontPlanIncludesAllHeaderFooterVariants()
    {
        var defaultHeader = CreateFontPlanParagraph(new DocxTextRun("DefaultHeader", 11d, null, false, false, false, null, "Default Sans")
        {
            Fonts = new DocxRunFonts("Default Sans", null, null, null, null, null, null, null)
        });
        var firstHeader = CreateFontPlanParagraph(new DocxTextRun("FirstHeader", 11d, null, false, false, false, null, "First Sans")
        {
            Fonts = new DocxRunFonts("First Sans", null, null, null, null, null, null, null)
        });
        var evenFooter = CreateFontPlanParagraph(new DocxTextRun("EvenFooter", 11d, null, false, false, false, null, "Even Sans")
        {
            Fonts = new DocxRunFonts("Even Sans", null, null, null, null, null, null, null)
        });
        var sectionHeader = CreateFontPlanParagraph(new DocxTextRun("SectionHeader", 11d, null, false, false, false, null, "Section Sans")
        {
            Fonts = new DocxRunFonts("Section Sans", null, null, null, null, null, null, null)
        });
        var sectionFooter = CreateFontPlanParagraph(new DocxTextRun("SectionFooter", 11d, null, false, false, false, null, "Section Footer Sans")
        {
            Fonts = new DocxRunFonts("Section Footer Sans", null, null, null, null, null, null, null)
        });
        DocxPageSettings sectionSettings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [sectionHeader]
            },
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [sectionFooter]
            }
        };
        var document = new DocxDocument(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [defaultHeader],
            [],
            [new DocxSectionBreakElement(sectionSettings, "nextPage", null, null, null, [])],
            [],
            [])
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [defaultHeader],
                ["first"] = [firstHeader]
            },
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["even"] = [evenFooter]
            }
        };
        var resolver = new MapFontResolver(["Default Sans", "First Sans", "Even Sans", "Section Sans", "Section Footer Sans"], "Resolver Fallback");

        string plannedTexts = string.Join("|", DocxFontPlan.Create(document, resolver).Runs.Select(run => run.Run.Text).Order(StringComparer.Ordinal));

        TestAssert.Equal("DefaultHeader|EvenFooter|FirstHeader|SectionFooter|SectionHeader", plannedTexts);
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

    public static void DocxParagraphLayoutStageMeasuresMixedRunSegmentsWithRunFontSizes()
    {
        var smallRun = new DocxTextRun("A", 10d, null, false, false, false, null, "Body");
        var largeRun = new DocxTextRun("B", 20d, null, false, false, false, null, "Body");
        var paragraph = new DocxParagraph(
            [smallRun, largeRun],
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
            .Create(document, new FontSizeWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal(20d, line.FontSize);
        TestAssert.Equal(10d, line.Segments[0].Width);
        TestAssert.Equal(20d, line.Segments[1].Width);
        TestAssert.Equal(10d, line.Segments[0].FontSize ?? 0d);
        TestAssert.Equal(20d, line.Segments[1].FontSize ?? 0d);
        TestAssert.Equal(line.Segments[0].X + line.Segments[0].Width, line.Segments[1].X);
    }

    public static void DocxParagraphLayoutStageAppliesVerticalAlignFontSizeAndBaseline()
    {
        var normalRun = new DocxTextRun("A", 16d, null, false, false, false, null, "Body");
        var superscriptRun = new DocxTextRun("2", 16d, null, false, false, false, null, "Body", VerticalAlignmentValue: "superscript");
        var subscriptRun = new DocxTextRun("n", 16d, null, false, false, false, null, "Body", VerticalAlignmentValue: "subscript");
        var paragraph = new DocxParagraph(
            [normalRun, superscriptRun, subscriptRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            16d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxDocument document = CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, new FontSizeWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal(16d, line.FontSize);
        TestAssert.Equal(16d, line.Segments[0].FontSize ?? 0d);
        TestAssert.Equal(10.5d, line.Segments[1].FontSize ?? 0d);
        TestAssert.Equal(10.5d, line.Segments[2].FontSize ?? 0d);
        TestAssert.Equal(0d, line.Segments[0].BaselineOffsetY);
        TestAssert.Equal(5.5d, line.Segments[1].BaselineOffsetY);
        TestAssert.Equal(-0.96d, line.Segments[2].BaselineOffsetY);
        TestAssert.Equal(line.Segments[0].X + line.Segments[0].Width, line.Segments[1].X);
        TestAssert.Equal(line.Segments[1].X + line.Segments[1].Width, line.Segments[2].X);
    }

    public static void DocxLayoutStageIncludesRunCharacterSpacingBetweenSegments()
    {
        var firstRun = new DocxTextRun("A", 12d, null, false, false, false, null, "Narrow", 3d);
        var secondRun = new DocxTextRun("B", 12d, null, false, false, false, null, "Narrow");
        var paragraph = new DocxParagraph(
            [firstRun, secondRun],
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
            75d,
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
            [],
            []);

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal(line.Segments[0].X + line.Segments[0].Width + 3d, line.Segments[1].X);
    }

    public static void DocxLayoutStageSplitsPreservedLeadingSpaceFromFollowingWord()
    {
        var firstRun = new DocxTextRun("A", 12d, null, false, false, false, null, "Narrow");
        var secondRun = new DocxTextRun(" B", 12d, null, false, false, false, null, "Wide");
        var paragraph = new DocxParagraph(
            [firstRun, secondRun],
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

        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal("A", line.Segments[0].Text);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal("B", line.Segments[2].Text);
        TestAssert.Equal(line.Segments[1].X + line.Segments[1].Width, line.Segments[2].X);
        TestAssert.Equal("Wide", line.Segments[1].StyleRun.FontFamily ?? string.Empty);
        TestAssert.Equal("Wide", line.Segments[2].StyleRun.FontFamily ?? string.Empty);
    }

    public static void DocxLayoutStageJustifiesWrappedNonFinalLines()
    {
        var run = new DocxTextRun("A B C D E F", 12d, null, false, false, false, null, "Narrow");
        var paragraph = new DocxParagraph(
            [run],
            [],
            null,
            DocxTextAlignment.Justified,
            "both",
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var document = new DocxDocument(
            72d,
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
            [],
            []);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        DocxTextLineLayout firstLine = lines[0];
        TestAssert.Equal(5, firstLine.Segments.Count);
        TestAssert.True(firstLine.Segments.All(segment => segment.Text.IndexOf(' ') < 0), "Expected justified DOCX layout to make word positions explicit instead of rendering stretchable spaces.");
        TestAssert.True(Math.Abs(firstLine.Width - 52d) < 0.001d, "Expected justified line inspection to expose the full paragraph width.");
        TestAssert.True(Math.Abs(firstLine.Segments.Last().X + firstLine.Segments.Last().Width - 62d) < 0.001d, "Expected non-final justified DOCX lines to stretch to the paragraph edge while excluding trailing wrap spaces.");
        TestAssert.True(lines[1].Width < firstLine.Width, "Expected the final line to keep its natural width instead of being justified.");
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

    public static void DocxParagraphLayoutStageWrapsMixedRunTextWithRunFontSizes()
    {
        var smallRun = new DocxTextRun("A ", 10d, null, false, false, false, null, "Body");
        var largeRun = new DocxTextRun("B", 20d, null, false, false, false, null, "Body");
        var paragraph = new DocxParagraph(
            [smallRun, largeRun],
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
            65d,
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
            .Create(document, new FontSizeWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal("A B", lines[0].Text);
        TestAssert.Equal(40d, lines[0].Width);
        TestAssert.Equal(2, lines[0].Segments.Count);
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

    public static void DocxRendererEmitsDistinctResourcesForResolvedRunTypefaces()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return;
        }

        (FontResolution Resolution, OpenTypeFont Font)? first = FindUsableInstalledFont();
        if (first is null)
        {
            return;
        }

        (FontResolution Resolution, OpenTypeFont Font)? second = FindUsableInstalledFontExcept(first.Value.Resolution.FamilyName);
        if (second is null ||
            string.Equals(first.Value.Resolution.FontFilePath, second.Value.Resolution.FontFilePath, StringComparison.OrdinalIgnoreCase) &&
            first.Value.Resolution.FontFaceIndex == second.Value.Resolution.FontFaceIndex)
        {
            return;
        }

        string firstFamily = System.Security.SecurityElement.Escape(first.Value.Resolution.FamilyName) ?? first.Value.Resolution.FamilyName;
        string secondFamily = System.Security.SecurityElement.Escape(second.Value.Resolution.FamilyName) ?? second.Value.Resolution.FamilyName;
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
                        <w:rPr><w:rFonts w:ascii="{{firstFamily}}" w:hAnsi="{{firstFamily}}"/></w:rPr>
                        <w:t>First face</w:t>
                      </w:r>
                      <w:r>
                        <w:rPr><w:rFonts w:ascii="{{secondFamily}}" w:hAnsi="{{secondFamily}}"/></w:rPr>
                        <w:t>Second face</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = new WindowsFontResolver(fontsDirectory) });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 ", pdf);
        TestAssert.Contains("/F2 ", pdf);
        TestAssert.Contains("/BaseFont /" + PdfEmbeddedFont.SanitizeName("LOKAD+" + first.Value.Resolution.FamilyName + "-"), pdf);
        TestAssert.Contains("/BaseFont /" + PdfEmbeddedFont.SanitizeName("LOKAD+" + second.Value.Resolution.FamilyName + "-"), pdf);
    }

    public static void DocxRendererDoesNotSynthesizeBoldWhenResolvedFaceIsBold()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
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
                        <w:rPr><w:b/><w:rFonts w:ascii="{{family}}" w:hAnsi="{{family}}"/></w:rPr>
                        <w:t>Resolved bold face</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var resolver = new SingleResolutionFontResolver(font.Value.Resolution);

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(2, CountPdfTextShows(pdf));
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

    public static void DocxReaderPreservesParagraphRunVerticalAlignmentTokens()
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
                  <w:style w:type="character" w:styleId="Raised">
                    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>Sub</w:t></w:r>
                      <w:r><w:rPr><w:rStyle w:val="Raised"/></w:rPr><w:t>Super</w:t></w:r>
                      <w:r><w:t>Base</w:t></w:r>
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

        TestAssert.Equal("subscript", runs[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("superscript", runs[1].VerticalAlignmentValue ?? string.Empty);
        TestAssert.True(runs[2].VerticalAlignmentValue is null, "Expected missing vertical alignment to keep a null source token.");
    }

    public static void DocxReaderPreservesParagraphRunStrikeTokens()
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
                  <w:style w:type="character" w:styleId="DoubleStrike">
                    <w:rPr><w:dstrike/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:strike/></w:rPr><w:t>Strike</w:t></w:r>
                      <w:r><w:rPr><w:strike w:val="0"/></w:rPr><w:t>NoStrike</w:t></w:r>
                      <w:r><w:rPr><w:rStyle w:val="DoubleStrike"/></w:rPr><w:t>Double</w:t></w:r>
                      <w:r><w:t>Plain</w:t></w:r>
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

        TestAssert.True(runs[0].Strike, "Expected w:strike to enable strike.");
        TestAssert.True(runs[0].StrikeValue is null, "Expected val-less w:strike to keep a null source token.");
        TestAssert.True(!runs[1].Strike, "Expected w:strike val=0 to disable strike.");
        TestAssert.Equal("0", runs[1].StrikeValue ?? string.Empty);
        TestAssert.True(runs[2].DoubleStrike, "Expected inherited w:dstrike to enable double strike.");
        TestAssert.True(runs[2].DoubleStrikeValue is null, "Expected val-less w:dstrike to keep a null source token.");
        TestAssert.True(!runs[3].Strike, "Expected missing strike to remain disabled.");
        TestAssert.True(runs[3].StrikeValue is null, "Expected missing strike to keep a null source token.");
        TestAssert.True(!runs[3].DoubleStrike, "Expected missing double strike to remain disabled.");
        TestAssert.True(runs[3].DoubleStrikeValue is null, "Expected missing double strike to keep a null source token.");
    }

    public static void DocxParagraphRendererDrawsTextDecorationsFromFontMetrics()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t>Under</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:strike/></w:rPr><w:t> Strike</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:dstrike/></w:rPr><w:t> Double</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int filledRectangles = pdf.Split(" re f", StringSplitOptions.None).Length - 1;
        TestAssert.True(filledRectangles >= 4, $"Expected underline, strike, and double-strike to render as filled metric rectangles; found {filledRectangles}.");
    }

    public static void DocxReaderPreservesParagraphRunHighlightAndShadingTokens()
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
                  <w:style w:type="character" w:styleId="Marked">
                    <w:rPr><w:highlight w:val="green"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Highlight</w:t></w:r>
                      <w:r><w:rPr><w:rStyle w:val="Marked"/></w:rPr><w:t>Inherited</w:t></w:r>
                      <w:r><w:rPr><w:shd w:val="pct20" w:color="112233" w:fill="D9EAD3"/></w:rPr><w:t>Shading</w:t></w:r>
                      <w:r><w:t>Plain</w:t></w:r>
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

        TestAssert.Equal("yellow", runs[0].HighlightValue ?? string.Empty);
        TestAssert.Equal("green", runs[1].HighlightValue ?? string.Empty);
        TestAssert.Equal("D9EAD3", runs[2].ShadingFillHex ?? string.Empty);
        TestAssert.Equal("pct20", runs[2].ShadingValue ?? string.Empty);
        TestAssert.Equal("112233", runs[2].ShadingColor ?? string.Empty);
        TestAssert.True(runs[3].HighlightValue is null, "Expected missing highlight to keep a null source token.");
        TestAssert.True(runs[3].ShadingFillHex is null, "Expected missing run shading fill to keep a null source token.");
        TestAssert.True(runs[3].ShadingValue is null, "Expected missing run shading value to keep a null source token.");
        TestAssert.True(runs[3].ShadingColor is null, "Expected missing run shading color to keep a null source token.");
    }

    public static void DocxParagraphRendererDrawsHighlightAndClearShadingBackgrounds()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/><w:highlight w:val="yellow"/></w:rPr><w:t>Yellow</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/><w:highlight w:val="darkBlue"/><w:color w:val="FFFFFF"/></w:rPr><w:t> Dark</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/><w:shd w:val="clear" w:fill="D9EAD3"/></w:rPr><w:t> Shade</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/><w:shd w:val="pct20" w:color="112233" w:fill="D9EAD3"/></w:rPr><w:t> PatternTokenOnly</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int filledRectangles = pdf.Split(" re f", StringSplitOptions.None).Length - 1;
        TestAssert.Equal(7, filledRectangles);
        TestAssert.Contains("1 1 0 rg", pdf);
        TestAssert.Contains("0 0 0.502 rg", pdf);
        TestAssert.Contains("0.851 0.918 0.827 rg", pdf);
        TestAssert.Contains("0.694 0.761 0.702 rg", pdf);
    }

    public static void DocxTableRendererBlendsPercentageCellShading()
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
                      <w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="2160"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/><w:shd w:val="pct20" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>Pattern</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:tcW w:w="2160" w:type="dxa"/><w:shd w:val="clear" w:fill="FCE5CD"/></w:tcPr><w:p><w:r><w:t>Clear</w:t></w:r></w:p></w:tc>
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
        TestAssert.Contains("0.694 0.761 0.702 rg", pdf);
        TestAssert.Contains("0.988 0.898 0.804 rg", pdf);
        TestAssert.DoesNotContain("0.851 0.918 0.827 rg", pdf);
    }

    public static void DocxReaderPreservesParagraphRunSmallCapsTokens()
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
                  <w:style w:type="character" w:styleId="SmallCapsStyle">
                    <w:rPr><w:smallCaps/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:smallCaps w:val="0"/></w:rPr><w:t>Off</w:t></w:r>
                      <w:r><w:rPr><w:rStyle w:val="SmallCapsStyle"/></w:rPr><w:t>Inherited</w:t></w:r>
                      <w:r><w:t>Plain</w:t></w:r>
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

        TestAssert.True(!runs[0].SmallCaps, "Expected w:smallCaps val=0 to disable small caps.");
        TestAssert.Equal("0", runs[0].SmallCapsValue ?? string.Empty);
        TestAssert.True(runs[1].SmallCaps, "Expected inherited w:smallCaps to enable small caps.");
        TestAssert.True(runs[1].SmallCapsValue is null, "Expected val-less inherited small caps to keep a null source token.");
        TestAssert.True(!runs[2].SmallCaps, "Expected missing small caps to remain disabled.");
        TestAssert.True(runs[2].SmallCapsValue is null, "Expected missing small caps to keep a null source token.");
    }

    public static void DocxParagraphLayoutSuppressesHiddenRunText()
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
                  <w:style w:type="character" w:styleId="HiddenStyle">
                    <w:rPr><w:vanish/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Visible</w:t></w:r>
                      <w:r><w:rPr><w:rStyle w:val="HiddenStyle"/></w:rPr><w:t>Hidden</w:t></w:r>
                      <w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>Shown</w:t></w:r>
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

        TestAssert.True(runs[1].Hidden, "Expected inherited w:vanish to mark the run as hidden.");
        TestAssert.True(runs[1].HiddenValue is null, "Expected val-less hidden style to keep a null source token.");
        TestAssert.True(!runs[2].Hidden, "Expected w:vanish val=0 to keep the run visible.");
        TestAssert.Equal("0", runs[2].HiddenValue ?? string.Empty);

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal("VisibleShown", line.Text);
        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("Visible", line.Segments[0].Text);
        TestAssert.Equal("Shown", line.Segments[1].Text);
    }

    public static void DocxReaderPreservesParagraphSimpleFieldCachedResultRunsInOrder()
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
                  <w:style w:type="character" w:styleId="ResultStyle">
                    <w:rPr><w:color w:val="336699"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:fldSimple w:instr=" DATE \@ &quot;yyyy&quot; ">
                        <w:r><w:rPr><w:rStyle w:val="ResultStyle"/></w:rPr><w:t>2026</w:t></w:r>
                      </w:fldSimple>
                      <w:r><w:t> After</w:t></w:r>
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

        TestAssert.Equal("Before ", runs[0].Text);
        TestAssert.Equal("2026", runs[1].Text);
        TestAssert.Equal("336699", runs[1].ColorHex ?? string.Empty);
        TestAssert.Equal(" After", runs[2].Text);
    }

    public static void DocxReaderPreservesFieldReferencesStructurally()
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
                      <w:fldSimple w:instr=" PAGE "/>
                      <w:fldSimple w:instr=" NUMPAGES "/>
                      <w:fldSimple w:instr=" PAGEREF Target "><w:r><w:t>TargetValue</w:t></w:r></w:fldSimple>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>1</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
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

        TestAssert.Equal(4, paragraph.FieldReferences.Count);
        TestAssert.Equal(2, paragraph.FieldReferences.Count(field => field.Kind == "Page"));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.Kind == "NumPages"));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.Kind == "Other"));
        TestAssert.Equal(3, paragraph.FieldReferences.Count(field => field.SourceKind == "Simple"));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.SourceKind == "ComplexInstruction"));
        TestAssert.Equal(2, paragraph.FieldReferences.Count(field => field.Placeholder == "{PAGE}"));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.Placeholder == "{NUMPAGES}"));
        TestAssert.True(paragraph.FieldReferences.Single(field => field.Instruction?.Contains("PAGEREF", StringComparison.Ordinal) == true).Placeholder is null, "PAGEREF should not be treated as a PAGE placeholder.");
        DocxFieldReference simplePage = paragraph.FieldReferences.First(field => field.Kind == "Page" && field.SourceKind == "Simple");
        DocxFieldReference simpleNumPages = paragraph.FieldReferences.Single(field => field.Kind == "NumPages");
        DocxFieldReference pageRef = paragraph.FieldReferences.Single(field => field.Kind == "Other");
        DocxFieldReference complexPage = paragraph.FieldReferences.Single(field => field.Kind == "Page" && field.SourceKind == "ComplexInstruction");
        TestAssert.Equal(0, simplePage.TextRunIndex);
        TestAssert.Equal(1, simplePage.TextRunCount);
        TestAssert.Equal(6, simplePage.TextLength);
        TestAssert.Equal(1, simpleNumPages.TextRunIndex);
        TestAssert.Equal(1, simpleNumPages.TextRunCount);
        TestAssert.Equal(10, simpleNumPages.TextLength);
        TestAssert.Equal(2, pageRef.TextRunIndex);
        TestAssert.Equal(1, pageRef.TextRunCount);
        TestAssert.Equal(11, pageRef.TextLength);
        TestAssert.Equal(3, complexPage.TextRunIndex);
        TestAssert.Equal(1, complexPage.TextRunCount);
        TestAssert.Equal(6, complexPage.TextLength);
        TestAssert.Equal("{PAGE}{NUMPAGES}TargetValue{PAGE}", string.Concat(paragraph.Runs.Select(run => run.Text)));

        DocxStructureSnapshot snapshot = new DocxRenderer().InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        DocxStructureStorySnapshot bodyStory = snapshot.Stories.Single(story => story.Kind == "Body");
        TestAssert.Equal(4, snapshot.FieldReferenceCount);
        TestAssert.Equal(2, snapshot.PageFieldReferenceCount);
        TestAssert.Equal(1, snapshot.NumPagesFieldReferenceCount);
        TestAssert.Equal(1, snapshot.OtherFieldReferenceCount);
        TestAssert.Equal(4, block.FieldReferenceCount);
        TestAssert.Equal(2, block.PageFieldReferenceCount);
        TestAssert.Equal(1, block.NumPagesFieldReferenceCount);
        TestAssert.Equal(1, block.OtherFieldReferenceCount);
        TestAssert.Equal(4, bodyStory.FieldReferenceCount);
    }

    public static void DocxComplexFieldWithCachedResultDoesNotEmitUnsupportedDiagnostic()
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
                      <w:r><w:t>Before </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE \@ &quot;yyyy&quot; </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>2026</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxTextRun[] runs = new DocxReader().Read(package).Paragraphs[0].Runs.ToArray();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before ", runs[0].Text);
        TestAssert.Equal("2026", runs[1].Text);
        TestAssert.Equal(" After", runs[2].Text);
    }

    public static void DocxComplexFieldCachedResultInsideHyperlinkDoesNotEmitUnsupportedDiagnostic()
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/field" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:hyperlink r:id="rIdLink">
                        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                        <w:r><w:instrText> DATE \@ &quot;yyyy&quot; </w:instrText></w:r>
                        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                        <w:r><w:t>2026</w:t></w:r>
                        <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      </w:hyperlink>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxParagraph paragraph = new DocxReader().Read(package).Paragraphs[0];

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before ", paragraph.Runs[0].Text);
        TestAssert.Equal("2026", paragraph.Runs[1].Text);
        TestAssert.Equal(" After", paragraph.Runs[2].Text);
        TestAssert.Equal(1, paragraph.Hyperlinks.Count);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].TextRunStartIndex);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].TextRunCount);
    }

    public static void DocxReaderUsesFinalViewForSimpleTrackedParagraphRuns()
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
                      <w:r><w:t>Before </w:t></w:r>
                      <w:ins><w:r><w:rPr><w:color w:val="336699"/></w:rPr><w:t>Inserted</w:t></w:r></w:ins>
                      <w:del><w:r><w:t>Deleted</w:t></w:r></w:del>
                      <w:r><w:t> After</w:t></w:r>
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

        TestAssert.Equal("Before ", runs[0].Text);
        TestAssert.Equal("Inserted", runs[1].Text);
        TestAssert.Equal("336699", runs[1].ColorHex ?? string.Empty);
        TestAssert.Equal(" After", runs[2].Text);
    }

    public static void DocxSimpleTrackedParagraphRunsDoNotEmitUnsupportedDiagnostic()
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
                      <w:r><w:t>Before </w:t></w:r>
                      <w:ins><w:r><w:t>Inserted</w:t></w:r></w:ins>
                      <w:del><w:r><w:t>Deleted</w:t></w:r></w:del>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_TRACKED_CHANGES", ids);
    }

    public static void DocxReaderPreservesParagraphHyperlinkRunsInOrder()
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
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" TargetMode="External"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="character" w:styleId="Hyperlink">
                    <w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:hyperlink r:id="rId2">
                        <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                      <w:r><w:t> After</w:t></w:r>
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
        DocxTextRun[] runs = paragraph.Runs.ToArray();

        TestAssert.Equal("Before ", runs[0].Text);
        TestAssert.Equal("Link", runs[1].Text);
        TestAssert.Equal("0563C1", runs[1].ColorHex ?? string.Empty);
        TestAssert.True(runs[1].Underline, "Expected hyperlink character style underline to survive.");
        TestAssert.Equal(" After", runs[2].Text);
        DocxHyperlinkSpan link = paragraph.Hyperlinks.Single();
        TestAssert.Equal("rId2", link.RelationshipId ?? string.Empty);
        TestAssert.Equal("https://example.invalid/", link.Target ?? string.Empty);
        TestAssert.Equal("External", link.TargetMode ?? string.Empty);
        TestAssert.True(link.ResolvedTarget is null, "External hyperlinks should keep their target without pretending to be package-local parts.");
        TestAssert.Equal(1, link.SourceRunStartIndex);
        TestAssert.Equal(1, link.SourceRunCount);
        TestAssert.Equal(1, link.TextRunStartIndex);
        TestAssert.Equal(1, link.TextRunCount);
        TestAssert.Equal(4, link.TextLength);

        DocxStructureSnapshot snapshot = new DocxRenderer().InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal(1, snapshot.HyperlinkCount);
        TestAssert.Equal(1, snapshot.ExternalHyperlinkCount);
        TestAssert.Equal(0, snapshot.InternalHyperlinkCount);
        TestAssert.Equal(1, block.HyperlinkCount);
        TestAssert.Equal(1, block.ExternalHyperlinkCount);
        TestAssert.Equal(0, block.InternalHyperlinkCount);
    }

    public static void DocxReaderPreservesInlineReferencesInsideRunContainers()
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
                      <w:r><w:t>Base</w:t></w:r>
                      <w:ins>
                        <w:r><w:t>Ins</w:t><w:footnoteReference w:id="5" w:customMarkFollows="1"/></w:r>
                      </w:ins>
                      <w:hyperlink w:anchor="Target">
                        <w:r><w:commentReference w:id="6"/><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                      <w:fldSimple w:instr=" REF Target ">
                        <w:r><w:t>Field</w:t><w:endnoteReference w:id="7"/></w:r>
                      </w:fldSimple>
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

        TestAssert.Equal(3, paragraph.InlineReferences.Count);
        DocxInlineReference footnote = paragraph.InlineReferences.Single(reference => reference.Kind == "Footnote");
        DocxInlineReference comment = paragraph.InlineReferences.Single(reference => reference.Kind == "Comment");
        DocxInlineReference endnote = paragraph.InlineReferences.Single(reference => reference.Kind == "Endnote");
        TestAssert.True(footnote.Id == "5" && footnote.CustomMarkFollowsValue == "1", "Inserted footnote marker metadata should survive run-container parsing.");
        TestAssert.True(comment.Id == "6" && comment.CustomMarkFollowsValue is null, "Hyperlink comment marker metadata should survive run-container parsing.");
        TestAssert.True(endnote.Id == "7" && endnote.CustomMarkFollowsValue is null, "Simple-field endnote marker metadata should survive run-container parsing.");
        TestAssert.Equal(1, footnote.SourceRunIndex);
        TestAssert.Equal(1, footnote.RunChildIndex);
        TestAssert.Equal(3, footnote.TextOffsetInRun);
        TestAssert.Equal(2, comment.SourceRunIndex);
        TestAssert.Equal(0, comment.RunChildIndex);
        TestAssert.Equal(0, comment.TextOffsetInRun);
        TestAssert.Equal(3, endnote.SourceRunIndex);
        TestAssert.Equal(1, endnote.RunChildIndex);
        TestAssert.Equal(5, endnote.TextOffsetInRun);

        DocxStructureSnapshot snapshot = new DocxRenderer().InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal(3, snapshot.InlineReferenceCount);
        TestAssert.Equal(3, snapshot.AnchoredInlineReferenceCount);
        TestAssert.Equal(1, block.CommentReferenceCount);
        TestAssert.Equal(1, block.FootnoteReferenceCount);
        TestAssert.Equal(1, block.EndnoteReferenceCount);
    }

    public static void DocxReaderPreservesBookmarkAnchorsForInternalHyperlinks()
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
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="7" w:name="Target"/>
                      <w:r><w:t>Target</w:t></w:r>
                      <w:bookmarkEnd w:id="7"/>
                      <w:hyperlink w:anchor="Target"><w:r><w:t>Jump</w:t></w:r></w:hyperlink>
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

        DocxBookmarkAnchor bookmark = paragraph.BookmarkAnchors.Single();
        TestAssert.Equal("7", bookmark.Id ?? string.Empty);
        TestAssert.Equal("Target", bookmark.Name ?? string.Empty);
        TestAssert.Equal(1, bookmark.SourceRunIndex);
        TestAssert.Equal(1, bookmark.TextRunIndex);
        TestAssert.Equal(7, bookmark.TextOffset);

        DocxHyperlinkSpan link = paragraph.Hyperlinks.Single();
        TestAssert.Equal("Target", link.Anchor ?? string.Empty);
        TestAssert.Equal(2, link.TextRunStartIndex);
        TestAssert.Equal(1, link.TextRunCount);

        DocxStructureSnapshot snapshot = new DocxRenderer().InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        DocxStructureStorySnapshot bodyStory = snapshot.Stories.Single(story => story.Kind == "Body");
        TestAssert.Equal(1, snapshot.BookmarkAnchorCount);
        TestAssert.Equal(1, block.BookmarkAnchorCount);
        TestAssert.Equal(1, bodyStory.BookmarkAnchorCount);
        TestAssert.Equal(1, snapshot.InternalHyperlinkCount);
    }

    public static void DocxReaderPreservesBookmarkAnchorsInsideHyperlinks()
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
                      <w:r><w:t>Before </w:t></w:r>
                      <w:hyperlink w:anchor="OuterTarget">
                        <w:bookmarkStart w:id="11" w:name="InnerTarget"/>
                        <w:r><w:t>Linked</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxParagraph paragraph = new DocxReader().Read(package).Paragraphs.Single();

        DocxBookmarkAnchor bookmark = paragraph.BookmarkAnchors.Single();
        TestAssert.Equal("InnerTarget", bookmark.Name ?? string.Empty);
        TestAssert.Equal(1, bookmark.SourceRunIndex);
        TestAssert.Equal(1, bookmark.TextRunIndex);
        TestAssert.Equal(7, bookmark.TextOffset);
        TestAssert.Equal("OuterTarget", paragraph.Hyperlinks.Single().Anchor ?? string.Empty);
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

    public static void DocxParagraphLayoutPreservesCarriageReturnsAsSoftLineBreaks()
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
                    <w:p><w:r><w:t>Alpha</w:t><w:cr/><w:t>Beta</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal("Alpha\nBeta", document.Paragraphs[0].Runs[0].Text);
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("Alpha", lines[0].Text);
        TestAssert.Equal("Beta", lines[1].Text);
    }

    public static void DocxParagraphLayoutPreservesTabsAsTabAdvances()
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
                    <w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal("A\tB", document.Paragraphs[0].Runs[0].Text);
        TestAssert.Equal("A\tB", line.Text);
        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("A", line.Segments[0].Text);
        TestAssert.Equal("B", line.Segments[1].Text);
        TestAssert.Equal(36d, line.Segments[1].X - line.X);
        TestAssert.Equal(41d, line.Width);
    }

    public static void DocxParagraphLayoutUsesAuthoredLeftTabStopsBeforeDefaultGrid()
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
                      <w:pPr><w:tabs><w:tab w:val="left" w:pos="1440" w:leader="dot"/></w:tabs></w:pPr>
                      <w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        DocxTabStop tabStop = document.Paragraphs[0].TabStops.Single();
        TestAssert.Equal(72d, tabStop.PositionPoints ?? 0d);
        TestAssert.Equal("1440", tabStop.PositionValue ?? string.Empty);
        TestAssert.Equal("left", tabStop.Value ?? string.Empty);
        TestAssert.Equal("dot", tabStop.LeaderValue ?? string.Empty);
        TestAssert.Equal(72d, line.Segments[1].X - line.X);
        TestAssert.Equal(77d, line.Width);
    }

    public static void DocxParagraphLayoutDoesNotUseBarOrClearTabsAsPositioningStops()
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
                        <w:tabs>
                          <w:tab w:val="bar" w:pos="360"/>
                          <w:tab w:val="clear" w:pos="720"/>
                        </w:tabs>
                      </w:pPr>
                      <w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, document.Paragraphs[0].TabStops.Count);
        TestAssert.Equal("bar", document.Paragraphs[0].TabStops[0].Value ?? string.Empty);
        TestAssert.Equal("clear", document.Paragraphs[0].TabStops[1].Value ?? string.Empty);
        TestAssert.Equal(36d, line.Segments[1].X - line.X);
        TestAssert.Equal(41d, line.Width);
    }

    public static void DocxReaderPreservesParagraphExplicitHyphenTokens()
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
                    <w:p><w:r><w:t>non</w:t><w:noBreakHyphen/><w:t>break</w:t><w:softHyphen/><w:t>soft</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        TestAssert.Equal("non\u2011break\u00ADsoft", document.Paragraphs[0].Runs[0].Text);
    }

    public static void DocxParagraphLayoutKeepsNonbreakingSpacesInsideWrapTokens()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("Alpha\u00A0Beta Gamma", 10d, null, false, false, false, null, null)],
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
            75d,
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
            [],
            []);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("Alpha\u00A0Beta ", lines[0].Text);
        TestAssert.Equal("Gamma", lines[1].Text);
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

    public static void DocxReaderPageBreakBeforeUsesResolvedParagraphStyleCascade()
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
                  <w:style w:type="paragraph" w:styleId="BreakBase">
                    <w:pPr><w:pageBreakBefore w:val="on"/></w:pPr>
                  </w:style>
                  <w:style w:type="paragraph" w:styleId="BreakChild">
                    <w:basedOn w:val="BreakBase"/>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>First</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pStyle w:val="BreakChild"/></w:pPr><w:r><w:t>Inherited break</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pStyle w:val="BreakChild"/><w:pageBreakBefore w:val="0"/></w:pPr><w:r><w:t>Override no break</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxPageBreakElement[] breaks = document.BodyElements.OfType<DocxPageBreakElement>().ToArray();

        TestAssert.Equal(1, breaks.Length);
        TestAssert.Equal("pageBreakBefore", breaks[0].SourceKind);
        TestAssert.Equal("on", breaks[0].Value ?? string.Empty);
        TestAssert.Equal(3, document.BodyElements.OfType<DocxParagraphElement>().Count());
    }

    public static void DocxReaderPromotesRunPageBreakOnlyParagraph()
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
                    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
                    <w:p><w:r><w:t>Second</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxBodyElement[] elements = document.BodyElements.ToArray();

        TestAssert.Equal(3, elements.Length);
        TestAssert.True(elements[1] is DocxPageBreakElement, "A run-level page-break-only paragraph should become a body page break.");
        var pageBreak = (DocxPageBreakElement)elements[1];
        TestAssert.Equal("runBreak", pageBreak.SourceKind);
        TestAssert.Equal("page", pageBreak.Value ?? string.Empty);
    }

    public static void DocxReaderPromotesRunColumnBreakOnlyParagraphAsManualBreak()
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
                    <w:p><w:r><w:br w:type="column"/></w:r></w:p>
                    <w:p><w:r><w:t>Second</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        DocxBodyElement[] elements = document.BodyElements.ToArray();

        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("First", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxManualBreakElement, "A run-level column-break-only paragraph should become a body manual break.");
        var manualBreak = (DocxManualBreakElement)elements[1];
        TestAssert.Equal("runBreak", manualBreak.SourceKind);
        TestAssert.Equal("column", manualBreak.Value ?? string.Empty);
        TestAssert.True(manualBreak.BreakParagraph is not null, "The authored break paragraph should remain available for future column-flow layout.");
        TestAssert.Equal("Second", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);

        DocxStructureSnapshot snapshot = new DocxRenderer().InspectStructure(document);
        TestAssert.Equal(1, snapshot.ManualBreakBlockCount);
        DocxStructureBlockSnapshot block = snapshot.Blocks[1];
        TestAssert.Equal("ManualBreak", block.Kind);
        TestAssert.Equal("Paragraph", block.PreviousKind ?? string.Empty);
        TestAssert.Equal("Paragraph", block.NextKind ?? string.Empty);
        TestAssert.Equal("runBreak", block.ManualBreakSourceKind ?? string.Empty);
        TestAssert.Equal("column", block.ManualBreakValue ?? string.Empty);
        TestAssert.True(block.ManualBreakConsumesParagraphLine == true, "The snapshot should expose the preserved break paragraph.");
    }

    public static void DocxReaderPromotesInlineRunPageBreakInsideParagraph()
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
                    <w:p><w:r><w:t>Alpha</w:t><w:br w:type="page"/><w:t>Beta</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxBodyElement[] elements = document.BodyElements.ToArray();
        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("Alpha", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxPageBreakElement, "The inline page break should become a body page break.");
        TestAssert.Equal("Beta", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("Alpha", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal("Beta", layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().Text);
    }

    public static void DocxReaderPromotesInlineRunColumnBreakInsideParagraph()
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
                    <w:p><w:r><w:t>Left</w:t><w:br w:type="column"/><w:t>Right</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="2"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_MANUAL_BREAK"), "Visible inline column breaks should lower to typed manual-break blocks, not stale unsupported diagnostics.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_MULTI_COLUMN"), "Explicit final-section column flow with authored column breaks should remain in the supported multi-column shape.");

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxBodyElement[] elements = document.BodyElements.ToArray();
        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("Left", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxManualBreakElement, "The inline column break should become a body manual break.");
        TestAssert.Equal("Right", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));
        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(2, snapshot.Pages[0].ColumnFrameCount);
        DocxLayoutItemSnapshot[] lines = snapshot.Pages[0].Items.Where(item => item.Kind == "TextLine").ToArray();
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(0, lines[0].ColumnIndex ?? -1);
        TestAssert.Equal(1, lines[1].ColumnIndex ?? -1);
    }

    public static void DocxReaderPromotesHyperlinkRunPageBreakInsideParagraph()
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
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:hyperlink r:id="rId2"><w:r><w:t>Alpha</w:t><w:br w:type="page"/><w:t>Beta</w:t></w:r></w:hyperlink></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxBodyElement[] elements = document.BodyElements.ToArray();
        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("Alpha", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxPageBreakElement, "The hyperlink-contained page break should become a body page break.");
        TestAssert.Equal("Beta", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);
    }

    public static void DocxReaderPromotesSimpleFieldRunPageBreakInsideParagraph()
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
                    <w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>Alpha</w:t><w:br w:type="page"/><w:t>Beta</w:t></w:r></w:fldSimple></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxBodyElement[] elements = document.BodyElements.ToArray();
        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("Alpha", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxPageBreakElement, "The simple-field-contained page break should become a body page break.");
        TestAssert.Equal("Beta", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);
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
                          <w:cols w:num="2" w:equalWidth="0" w:space="720">
                            <w:col w:w="3000" w:space="360"/>
                            <w:col w:w="4200"/>
                          </w:cols>
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
        TestAssert.Equal(2, sectionBreak.ColumnDefinitions.Count);
        TestAssert.Equal("3000", sectionBreak.ColumnDefinitions[0].WidthValue ?? string.Empty);
        TestAssert.Equal("360", sectionBreak.ColumnDefinitions[0].SpaceValue ?? string.Empty);
        TestAssert.Equal("4200", sectionBreak.ColumnDefinitions[1].WidthValue ?? string.Empty);
        TestAssert.True(sectionBreak.ColumnDefinitions[1].SpaceValue is null, "The final custom column should not invent a trailing gutter token.");
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Paragraph section break should remain anchored after its paragraph.");
        TestAssert.True(document.BodyElements[1] is DocxSectionBreakElement, "Section break should be part of body flow.");

        DocxStructureSnapshot structure = DocxStructureSnapshot.FromDocument(document);
        TestAssert.Equal(1, structure.SectionBreakBlockCount);
        TestAssert.Equal(1, structure.ContinuousSectionBreakBlockCount);
        TestAssert.Equal(0, structure.PageStartingSectionBreakBlockCount);
        TestAssert.Equal(0, structure.DefaultSectionBreakBlockCount);
        TestAssert.Equal(1, structure.ColumnSectionBreakBlockCount);

        DocxStructureBlockSnapshot structureSnapshot = structure.Blocks[1];
        TestAssert.Equal(2, structureSnapshot.SectionColumnDefinitionCount ?? 0);
        TestAssert.Equal(2, structureSnapshot.SectionColumnDefinitionWidthTokenCount ?? 0);
        TestAssert.Equal(1, structureSnapshot.SectionColumnDefinitionSpaceTokenCount ?? 0);
    }

    public static void DocxSectionBreakNextPageStartsNewLayoutPage()
    {
        var first = new DocxParagraph(
            [new DocxTextRun("First", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.25d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var second = first with
        {
            Runs = [new DocxTextRun("Second", 11d, null, false, false, false, null, null)]
        };
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(first),
                new DocxSectionBreakElement(DocxPageSettings.Empty, "nextPage", null, null, null, []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("First", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal("Second", layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().Text);
    }

    public static void DocxSectionBreakPageSettingsOwnPrecedingLayoutSection()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph second = CreateDocxLayoutParagraph("Second", fontSize: 10d, lineSpacingPoints: 10d);
        var firstSectionSettings = new DocxPageSettings(
            "4000",
            "4000",
            null,
            "360",
            "360",
            "360",
            "360",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        var finalSectionSettings = new DocxPageSettings(
            "6000",
            "6000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        var document = new DocxDocument(
            300d,
            300d,
            72d,
            72d,
            72d,
            72d,
            finalSectionSettings,
            [],
            [],
            [],
            [
                new DocxParagraphElement(first),
                new DocxSectionBreakElement(firstSectionSettings, "nextPage", "2", "1", "720", []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(200d, layout.Pages[0].Width);
        TestAssert.Equal(200d, layout.Pages[0].Height);
        TestAssert.Equal(18d, layout.Pages[0].MarginLeft);
        TestAssert.Equal(18d, layout.Pages[0].MarginTop);
        TestAssert.Equal("360", layout.Pages[0].PageSettings.MarginLeftValue ?? string.Empty);
        TestAssert.Equal("nextPage", layout.Pages[0].SectionProperties.BreakTypeValue ?? string.Empty);
        TestAssert.Equal("2", layout.Pages[0].SectionProperties.ColumnCountValue ?? string.Empty);
        TestAssert.Equal("1", layout.Pages[0].SectionProperties.ColumnEqualWidthValue ?? string.Empty);
        TestAssert.Equal("720", layout.Pages[0].SectionProperties.ColumnSpaceValue ?? string.Empty);
        TestAssert.Equal(2, layout.Pages[0].SectionProperties.ColumnCount ?? 0);
        TestAssert.Equal(36d, layout.Pages[0].SectionProperties.ColumnSpacePoints ?? 0d);
        TestAssert.Equal(2, layout.Pages[0].ColumnFrames.Count);
        TestAssert.Equal(18d, layout.Pages[0].ColumnFrames[0].X);
        TestAssert.Equal(64d, layout.Pages[0].ColumnFrames[0].Width);
        TestAssert.Equal(36d, layout.Pages[0].ColumnFrames[0].GutterAfterPoints ?? 0d);
        TestAssert.Equal(118d, layout.Pages[0].ColumnFrames[1].X);
        TestAssert.Equal(64d, layout.Pages[0].ColumnFrames[1].Width);
        TestAssert.Equal(18d, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().X);
        TestAssert.Equal(300d, layout.Pages[1].Width);
        TestAssert.Equal(300d, layout.Pages[1].Height);
        TestAssert.Equal(72d, layout.Pages[1].MarginLeft);
        TestAssert.Equal(72d, layout.Pages[1].MarginTop);
        TestAssert.Equal("1440", layout.Pages[1].PageSettings.MarginLeftValue ?? string.Empty);
        TestAssert.True(layout.Pages[1].SectionProperties.ColumnCountValue is null, "Final section page should not inherit previous section column tokens.");
        TestAssert.Equal(1, layout.Pages[1].ColumnFrames.Count);
        TestAssert.Equal(72d, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().X);

        DocxLayoutPageSnapshot firstPageSnapshot = DocxLayoutSnapshot.FromLayout(layout).Pages[0];
        TestAssert.Equal(2, firstPageSnapshot.ColumnFrameCount);
        TestAssert.Equal(128d, firstPageSnapshot.ColumnFrameWidthSum);
        TestAssert.Equal(36d, firstPageSnapshot.ColumnGutterWidthSum);
        TestAssert.Equal(118d, firstPageSnapshot.ColumnFrames[1].X);
    }

    public static void DocxContinuousSectionBreakOnEmptyPageAppliesFollowingSectionGeometry()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph second = CreateDocxLayoutParagraph("Second", fontSize: 10d, lineSpacingPoints: 10d);
        var firstSectionSettings = new DocxPageSettings(
            "4000",
            "4000",
            null,
            "360",
            "360",
            "360",
            "360",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        var finalSectionSettings = new DocxPageSettings(
            "6000",
            "6000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        var document = new DocxDocument(
            300d,
            300d,
            72d,
            72d,
            72d,
            72d,
            finalSectionSettings,
            [],
            [],
            [],
            [
                new DocxParagraphElement(first),
                new DocxPageBreakElement("runBreak", "page"),
                new DocxSectionBreakElement(firstSectionSettings, "continuous", null, null, null, []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(200d, layout.Pages[0].Width);
        TestAssert.Equal(18d, layout.Pages[0].MarginLeft);
        TestAssert.Equal("continuous", layout.Pages[0].SectionProperties.BreakTypeValue ?? string.Empty);
        TestAssert.Equal(300d, layout.Pages[1].Width);
        TestAssert.Equal(72d, layout.Pages[1].MarginLeft);
        TestAssert.True(layout.Pages[1].SectionProperties.BreakTypeValue is null, "Following final section geometry should replace the continuous break on an empty page.");
        TestAssert.Equal(72d, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().X);
    }

    public static void DocxOddPageSectionBreakInsertsBlankParityPage()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph second = CreateDocxLayoutParagraph("Second", fontSize: 10d, lineSpacingPoints: 10d);
        var firstSectionSettings = new DocxPageSettings(
            "4000",
            "4000",
            null,
            "360",
            "360",
            "360",
            "360",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        var finalSectionSettings = new DocxPageSettings(
            "6000",
            "6000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        var document = new DocxDocument(
            300d,
            300d,
            72d,
            72d,
            72d,
            72d,
            finalSectionSettings,
            [],
            [],
            [],
            [
                new DocxParagraphElement(first),
                new DocxSectionBreakElement(firstSectionSettings, "oddPage", null, null, null, []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal("First", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal(0, layout.Pages[1].Items.Count);
        TestAssert.Equal(200d, layout.Pages[1].Width);
        TestAssert.Equal("oddPage", layout.Pages[1].SectionProperties.BreakTypeValue ?? string.Empty);
        TestAssert.Equal("Second", layout.Pages[2].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal(300d, layout.Pages[2].Width);
        TestAssert.True(layout.Pages[2].SectionProperties.BreakTypeValue is null, "Following section should start after the inserted parity page.");
    }

    public static void DocxSectionBreakCustomColumnsCreatePageOwnedFrames()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 12d);
        DocxParagraph second = first with
        {
            Runs = [new DocxTextRun("Second", 10d, null, false, false, false, null, null)]
        };
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            WidthValue = "6000",
            HeightValue = "6000",
            MarginLeftValue = "360",
            MarginRightValue = "360",
            MarginTopValue = "360",
            MarginBottomValue = "360"
        };
        var document = new DocxDocument(
            300d,
            300d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(first),
                new DocxSectionBreakElement(
                    firstSectionSettings,
                    "nextPage",
                    "2",
                    "0",
                    "720",
                    [
                        new DocxSectionColumn("2000", "240"),
                        new DocxSectionColumn("2400", null)
                    ]),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(2, layout.Pages[0].SectionProperties.ColumnDefinitions.Count);
        TestAssert.Equal(100d, layout.Pages[0].SectionProperties.ColumnDefinitions[0].WidthPoints ?? 0d);
        TestAssert.Equal(12d, layout.Pages[0].SectionProperties.ColumnDefinitions[0].SpacePoints ?? 0d);
        TestAssert.Equal(2, layout.Pages[0].ColumnFrames.Count);
        TestAssert.Equal(18d, layout.Pages[0].ColumnFrames[0].X);
        TestAssert.Equal(100d, layout.Pages[0].ColumnFrames[0].Width);
        TestAssert.Equal(12d, layout.Pages[0].ColumnFrames[0].GutterAfterPoints ?? 0d);
        TestAssert.Equal(130d, layout.Pages[0].ColumnFrames[1].X);
        TestAssert.Equal(120d, layout.Pages[0].ColumnFrames[1].Width);

        DocxLayoutPageSnapshot firstPageSnapshot = DocxLayoutSnapshot.FromLayout(layout).Pages[0];
        TestAssert.Equal(2, firstPageSnapshot.SectionColumnDefinitionCount);
        TestAssert.Equal(220d, firstPageSnapshot.SectionColumnDefinitionWidthSum);
        TestAssert.Equal(12d, firstPageSnapshot.SectionColumnDefinitionSpaceSum);
        TestAssert.Equal(220d, firstPageSnapshot.ColumnFrameWidthSum);
        TestAssert.Equal(12d, firstPageSnapshot.ColumnGutterWidthSum);
    }

    public static void DocxPageLayoutStageManualColumnBreakAdvancesActiveColumn()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 12d);
        DocxParagraph second = first with
        {
            Runs = [new DocxTextRun("Second", 10d, null, false, false, false, null, null)]
        };
        DocxPageSettings sectionSettings = DocxPageSettings.Empty with
        {
            WidthValue = "4000",
            HeightValue = "4000",
            MarginLeftValue = "360",
            MarginRightValue = "360",
            MarginTopValue = "360",
            MarginBottomValue = "360"
        };
        var anchoredDrawing = new DocxFloatingDrawing(
            "0",
            "0",
            "0",
            "0",
            "0",
            "0",
            "0",
            "0",
            "1",
            "1",
            "914400",
            "457200",
            "column",
            "left",
            null,
            "paragraph",
            null,
            "0",
            "square",
            "bothSides",
            SourceParagraphIndex: 0,
            SourceBlockIndex: 2);
        var document = new DocxDocument(
            300d,
            300d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [anchoredDrawing],
            [],
            [],
            [
                new DocxParagraphElement(first),
                new DocxManualBreakElement("runBreak", "column"),
                new DocxParagraphElement(second),
                new DocxSectionBreakElement(sectionSettings, "nextPage", "2", "1", "360", [])
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTextLineLayout[] lines = layout.Pages[0].Items.OfType<DocxTextLineLayout>().ToArray();

        TestAssert.Equal(1, layout.Pages.Count);
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(18d, lines[0].X);
        TestAssert.Equal(109d, lines[1].X);
        TestAssert.Equal(2, lines[1].SourceBlockIndex ?? -1);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(0, snapshot.Pages[0].Items[0].ColumnIndex ?? -1);
        TestAssert.Equal(1, snapshot.Pages[0].Items[1].ColumnIndex ?? -1);
        TestAssert.Equal(0, snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 0).FirstColumnIndex ?? -1);
        TestAssert.Equal(1, snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 2).FirstColumnIndex ?? -1);
        DocxFloatingDrawingLayoutSnapshot anchorSnapshot = snapshot.FloatingDrawings.Single();
        TestAssert.Equal(1, anchorSnapshot.AnchorColumnIndex ?? -1);
        TestAssert.Equal(109d, anchorSnapshot.HorizontalReferenceX ?? 0d);
        TestAssert.Equal(73d, anchorSnapshot.HorizontalReferenceWidth ?? 0d);
        TestAssert.Equal(109d, anchorSnapshot.PlacedX ?? 0d);
        TestAssert.Equal(anchorSnapshot.VerticalReferenceTop ?? 0d, anchorSnapshot.PlacedTop ?? -1d);
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

    public static void DocxLayoutStageAppliesParagraphIndentsToWrapping()
    {
        DocxParagraph indented = CreateDocxLayoutParagraph(
            "One two three four five six seven eight nine",
            fontSize: 10d,
            lineSpacingPoints: 10d,
            indent: new DocxParagraphIndent(20d, 10d, 10d, null, "400", "200", "200", null));
        var document = new DocxDocument(
            120d,
            120d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(indented)],
            [indented],
            []);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages.Single()
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.True(lines.Length > 1, "Indented paragraph should wrap so first and continuation indents are both exercised.");
        TestAssert.Equal(40d, lines[0].X);
        TestAssert.Equal(30d, lines[1].X);
        TestAssert.True(lines[0].Width <= 60d, "First-line width should subtract left, first-line, and right indents.");
        TestAssert.True(lines[1].Width <= 70d, "Continuation width should subtract left and right indents.");
    }

    public static void DocxLayoutStageUsesFontMetricsForAutoLineHeight()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        var first = new DocxParagraph(
            [new DocxTextRun("First", 10d, null, false, false, false, null, font.Value.Resolution.FamilyName)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.15d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var second = first with
        {
            Runs = [new DocxTextRun("Second", 10d, null, false, false, false, null, font.Value.Resolution.FamilyName)]
        };
        var body = new DocxBodyElement[] { new DocxParagraphElement(first), new DocxParagraphElement(second) };
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
            body,
            [first, second],
            []);
        var resolver = new SingleResolutionFontResolver(font.Value.Resolution);
        DocxFontPlan plan = DocxFontPlan.Create(document, resolver);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new DocxFontPlanTextMeasurer(plan))
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        double typographicUnits = font.Value.Font.Os2.TypographicAscender -
            font.Value.Font.Os2.TypographicDescender +
            font.Value.Font.Os2.TypographicLineGap;
        double expectedSingleLineHeight = Math.Max(11.5d, typographicUnits * 10d / font.Value.Font.UnitsPerEm);
        double expected = expectedSingleLineHeight * 1.15d;
        double actual = lines[0].BaselineY - lines[1].BaselineY;
        TestAssert.True(Math.Abs(actual - expected) < 0.01d, $"Auto DOCX line height should advance on the resolved typographic font line box, not the em size or Windows bounding box. Expected {expected}, actual {actual}.");
    }

    public static void DocxLayoutStageConsumesEmptyParagraphLineBox()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("Alpha", 10d, 10d);
        var empty = new DocxParagraph(
            [],
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
            null);
        DocxParagraph second = CreateDocxLayoutParagraph("Beta", 10d, 10d);
        DocxDocument document = CreateLayoutTestDocument(
            [new DocxParagraphElement(first), new DocxParagraphElement(empty), new DocxParagraphElement(second)],
            []);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("Alpha", lines[0].Text);
        TestAssert.Equal("Beta", lines[1].Text);
        TestAssert.Equal(20d, Math.Round(lines[0].BaselineY - lines[1].BaselineY, 3));
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
        TestAssert.Equal(0d, lines[0].PendingAfterSpacing ?? -1d);
        TestAssert.Equal(10d, lines[0].ParagraphBeforeSpacing ?? -1d);
        TestAssert.Equal(10d, lines[0].ParagraphAfterSpacing ?? -1d);
        TestAssert.Equal(10d, lines[0].AppliedBeforeSpacing ?? -1d);
        TestAssert.True(lines[0].ContextualSpacingSuppressed == false, "The first same-style paragraph should expose that contextual spacing did not suppress its boundary.");
        TestAssert.Equal(10d, lines[1].PendingAfterSpacing ?? -1d);
        TestAssert.Equal(10d, lines[1].ParagraphBeforeSpacing ?? -1d);
        TestAssert.Equal(10d, lines[1].ParagraphAfterSpacing ?? -1d);
        TestAssert.Equal(0d, lines[1].AppliedBeforeSpacing ?? -1d);
        TestAssert.True(lines[1].ContextualSpacingSuppressed == true, "The second same-style paragraph should expose the contextual spacing suppression decision.");
    }

    public static void DocxSyntheticContextualSpacingKeepsDifferentStyleGap()
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
            "BodyA",
            DocxTextAlignment.Left,
            null,
            10d,
            10d,
            1d,
            10d,
            spacing,
            DocxParagraphKeepRules.Empty,
            null);
        var second = first with
        {
            Runs = [new DocxTextRun("Second", 10d, null, false, false, false, null, null)],
            StyleId = "BodyB"
        };
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
        TestAssert.Equal(20d, Math.Round(lines[0].BaselineY - lines[1].BaselineY, 3));
        TestAssert.Equal(10d, lines[1].PendingAfterSpacing ?? -1d);
        TestAssert.Equal(10d, lines[1].ParagraphBeforeSpacing ?? -1d);
        TestAssert.Equal(10d, lines[1].AppliedBeforeSpacing ?? -1d);
        TestAssert.True(lines[1].ContextualSpacingSuppressed == false, "Different styles should keep the authored boundary gap even when contextual spacing is enabled.");
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

    public static void DocxSyntheticPageKeepNextEstimatesIndentedBlockTarget()
    {
        DocxParagraph keepNext = CreateDocxLayoutParagraph(
            "Keep",
            fontSize: 10d,
            lineSpacingPoints: 10d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph tableParagraph = CreateDocxLayoutParagraph("One Two Three Four Five", fontSize: 10d, lineSpacingPoints: 10d);
        var table = new DocxTable(
            null,
            [100d],
            [new DocxTableRow([new DocxTableCell("One Two Three Four Five", [tableParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 10d)],
            PreferredWidthPoints: 100d,
            IndentPoints: 70d);
        DocxParagraph[] fillers =
        [
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
        ];
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Concat([new DocxParagraphElement(keepNext), new DocxTableElement(table)])
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
            [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTextLineLayout[] firstPageLines = layout.Pages[0].Items.OfType<DocxTextLineLayout>().ToArray();
        DocxTextLineLayout[] secondPageLines = layout.Pages[1].Items.OfType<DocxTextLineLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(3, firstPageLines.Length);
        TestAssert.Equal("Keep", secondPageLines[0].Text);
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
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

    public static void DocxSyntheticParagraphDefaultWidowControlMovesThreeLineParagraphToNextPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        DocxParagraph widowControlled = CreateDocxLayoutParagraph("One\nTwo\nThree", fontSize: 10d, lineSpacingPoints: 10d);
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

        DocxTextLineLayout[] secondPageLines = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[1]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(3, secondPageLines.Length);
        TestAssert.Equal("One", secondPageLines[0].Text);
    }

    public static void DocxSyntheticParagraphExplicitWidowControlOffAllowsWidowLine()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        DocxParagraph widowOff = CreateDocxLayoutParagraph(
            "One\nTwo\nThree",
            fontSize: 10d,
            lineSpacingPoints: 10d,
            keepRules: new DocxParagraphKeepRules(null, null, null, null, false, null));
        DocxParagraph[] fillers =
        [
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
        ];
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Append(new DocxParagraphElement(widowOff))
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

        TestAssert.Equal(6, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Count());
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
        TestAssert.Equal(2, Regex.Matches(pdf, "0\\.044 Tc").Count);
        TestAssert.Equal(8, CountPdfTextShows(pdf));
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxListLabel? label = document.Paragraphs.Single().ListLabel;
        TestAssert.Equal("\uF0A7", label?.Text ?? string.Empty);
        TestAssert.Equal("Marker Symbols", label?.Style.Fonts.Ascii ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
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
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Indented".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout line = new DocxLayoutEngine()
            .Create(document, embedded)
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
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal(expectedTextX, line.Segments[2].X);
        TestAssert.True(line.Segments[2].X < 108d, "A space suffix should not advance text to the numbering tab stop.");
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "1. Alpha Beta".EnumerateRunes().Select(rune => rune.Value));

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, embedded)
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
                          <wp:anchor distT="114300" distB="228600" distL="342900" distR="457200"
                                     simplePos="0" relativeHeight="251658240" behindDoc="0"
                                     locked="1" layoutInCell="1" allowOverlap="0">
                            <wp:extent cx="1828800" cy="914400"/>
                            <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:posOffset>63500</wp:posOffset></wp:positionV>
                            <wp:wrapSquare wrapText="bothSides"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
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
        TestAssert.Equal("rIdImage1", drawing.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/image1.png", drawing.Image?.PartName ?? string.Empty);
        TestAssert.Equal("image/png", drawing.Image?.ContentType ?? string.Empty);
        TestAssert.Equal(144d, drawing.Image?.WidthPoints ?? 0d);
        TestAssert.Equal(72d, drawing.Image?.HeightPoints ?? 0d);
        TestAssert.Equal(0, drawing.SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, drawing.SourceBlockIndex ?? -1);

        DocxStructureFloatingDrawingSnapshot snapshot = new DocxRenderer().InspectStructure(document).FloatingDrawings.Single();
        TestAssert.Equal("rIdImage1", snapshot.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/image1.png", snapshot.ImagePartName ?? string.Empty);
        TestAssert.Equal("image/png", snapshot.ImageContentType ?? string.Empty);
        TestAssert.Equal(144d, snapshot.ImageWidthPoints ?? 0d);
        TestAssert.Equal(72d, snapshot.ImageHeightPoints ?? 0d);
        TestAssert.Equal(0, snapshot.SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, snapshot.SourceBlockIndex ?? -1);

        DocxFloatingDrawingLayoutSnapshot layoutSnapshot = new DocxRenderer().InspectLayout(document).FloatingDrawings.Single();
        TestAssert.Equal(0, layoutSnapshot.SourceBlockIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.PageStartIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.PageEndIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.AnchorPageIndex ?? -1);
        TestAssert.True(layoutSnapshot.AnchorBlockVerticalTop is not null, "Floating drawing layout should resolve its source block top before rendering.");
        TestAssert.True(layoutSnapshot.AnchorBlockVerticalBottom is not null, "Floating drawing layout should resolve its source block bottom before rendering.");
        TestAssert.True(layoutSnapshot.AnchorBlockVerticalTop > layoutSnapshot.AnchorBlockVerticalBottom, "Floating drawing anchor block should carry placed vertical bounds.");
        TestAssert.Equal(144d, layoutSnapshot.ExtentWidthPoints ?? 0d);
        TestAssert.Equal(72d, layoutSnapshot.ExtentHeightPoints ?? 0d);
        TestAssert.True(layoutSnapshot.HorizontalOffsetPoints is null, "Aligned horizontal anchors should not invent a numeric offset.");
        TestAssert.Equal(5d, layoutSnapshot.VerticalOffsetPoints ?? 0d);
        TestAssert.Equal(9d, layoutSnapshot.DistanceTopPoints ?? 0d);
        TestAssert.Equal(18d, layoutSnapshot.DistanceBottomPoints ?? 0d);
        TestAssert.Equal(27d, layoutSnapshot.DistanceLeftPoints ?? 0d);
        TestAssert.Equal(36d, layoutSnapshot.DistanceRightPoints ?? 0d);
        DocxLayoutPageSnapshot layoutPage = new DocxRenderer().InspectLayout(document).Pages.Single();
        TestAssert.Equal(layoutPage.MarginLeft, layoutSnapshot.HorizontalReferenceX ?? 0d);
        TestAssert.Equal(layoutPage.Width - layoutPage.MarginLeft - layoutPage.MarginRight, layoutSnapshot.HorizontalReferenceWidth ?? 0d);
        TestAssert.Equal(layoutSnapshot.AnchorBlockVerticalTop ?? 0d, layoutSnapshot.VerticalReferenceTop ?? 0d);
        TestAssert.Equal(layoutSnapshot.AnchorBlockVerticalBottom ?? 0d, layoutSnapshot.VerticalReferenceBottom ?? 0d);
        TestAssert.Equal("Align", layoutSnapshot.HorizontalPlacementSource ?? string.Empty);
        TestAssert.Equal("Offset", layoutSnapshot.VerticalPlacementSource ?? string.Empty);
        TestAssert.Equal("rIdImage1", layoutSnapshot.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/image1.png", layoutSnapshot.ImagePartName ?? string.Empty);
        TestAssert.Equal("image/png", layoutSnapshot.ImageContentType ?? string.Empty);
        TestAssert.Equal(144d, layoutSnapshot.ImageWidthPoints ?? 0d);
        TestAssert.Equal(72d, layoutSnapshot.ImageHeightPoints ?? 0d);
        TestAssert.Equal((layoutSnapshot.PlacedX ?? 0d) - (layoutSnapshot.DistanceLeftPoints ?? 0d), layoutSnapshot.WrapExclusionX ?? 0d);
        TestAssert.Equal((layoutSnapshot.PlacedTop ?? 0d) + (layoutSnapshot.DistanceTopPoints ?? 0d), layoutSnapshot.WrapExclusionTop ?? 0d);
        TestAssert.Equal((layoutSnapshot.ExtentWidthPoints ?? 0d) + (layoutSnapshot.DistanceLeftPoints ?? 0d) + (layoutSnapshot.DistanceRightPoints ?? 0d), layoutSnapshot.WrapExclusionWidth ?? 0d);
        TestAssert.Equal((layoutSnapshot.ExtentHeightPoints ?? 0d) + (layoutSnapshot.DistanceTopPoints ?? 0d) + (layoutSnapshot.DistanceBottomPoints ?? 0d), layoutSnapshot.WrapExclusionHeight ?? 0d);

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.True(
            !diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_FLOATING_DRAWING"),
            "Structurally supported rendered floating image anchors should not emit stale unsupported-floating diagnostics.");
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        double placedImageY = (layoutSnapshot.PlacedTop ?? 0d) - (layoutSnapshot.ExtentHeightPoints ?? 0d);
        string imageTransform = string.Create(
            CultureInfo.InvariantCulture,
            $"{layoutSnapshot.ExtentWidthPoints ?? 0d:0.###} 0 0 {layoutSnapshot.ExtentHeightPoints ?? 0d:0.###} {layoutSnapshot.PlacedX ?? 0d:0.###} {placedImageY:0.###} cm");
        TestAssert.Contains(imageTransform, pdf);
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
                      <w:tblPr>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideH w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideV w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
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
        TestAssert.Contains(" re f", pdf);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.True(CountPdfTextShows(pdf) >= 1, "Expected DOCX paragraph text to render as a PDF text-show operation.");
    }

    public static void DocxSyntheticTableWithoutBordersDoesNotInventCellGrid()
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
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>No border</w:t></w:r></w:p></w:tc>
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
        TestAssert.DoesNotContain(" re S", pdf);
        TestAssert.True(CountPdfTextShows(pdf) >= 1, "Expected DOCX table text to render as a PDF text-show operation.");
    }

    public static void DocxSyntheticTableTextRendersWithRunFontResourceWithoutFallback()
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
                      <w:tblPr><w:tblW w:w="2880" w:type="dxa"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:p>
                            <w:r>
                              <w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="22"/></w:rPr>
                              <w:t>Table explicit font</w:t>
                            </w:r>
                          </w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
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
        TestAssert.Contains("/F1 11.04 Tf", pdf);
        TestAssert.True(CountPdfTextShows(pdf) >= 1, "Expected DOCX table text to render as a PDF text-show operation.");
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

    public static void DocxReaderAppliesDefaultTableStyleCellMargins()
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
                  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
                    <w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
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
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));
        DocxTableCellMargins margins = document.Tables[0].Rows[0].Cells[0].Margins;

        TestAssert.Equal("108", margins.LeftValue ?? string.Empty);
        TestAssert.Equal(5.4d, margins.LeftPoints ?? 0d);
        TestAssert.Equal("120", margins.RightValue ?? string.Empty);
        TestAssert.Equal(6d, margins.RightPoints ?? 0d);
    }

    public static void DocxReaderDirectTableCellMarginsOverrideDefaultTableStyle()
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
                  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
                    <w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblCellMar><w:left w:w="240" w:type="dxa"/></w:tblCellMar></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>Direct</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));
        DocxTableCellMargins margins = document.Tables[0].Rows[0].Cells[0].Margins;

        TestAssert.Equal("240", margins.LeftValue ?? string.Empty);
        TestAssert.Equal(12d, margins.LeftPoints ?? 0d);
        TestAssert.Equal("108", margins.RightValue ?? string.Empty);
        TestAssert.Equal(5.4d, margins.RightPoints ?? 0d);
    }

    public static void DocxReaderTableStyleCascadesBasedOnProperties()
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
                  <w:style w:type="table" w:styleId="BaseTable">
                    <w:tblPr><w:tblCellMar><w:left w:w="240" w:type="dxa"/></w:tblCellMar></w:tblPr>
                    <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9EAD3"/></w:tcPr>
                    <w:pPr><w:jc w:val="right"/></w:pPr>
                    <w:rPr><w:b/><w:color w:val="336699"/></w:rPr>
                    <w:tblStylePr w:type="firstRow"><w:tcPr><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr>
                  </w:style>
                  <w:style w:type="table" w:styleId="ChildTable">
                    <w:basedOn w:val="BaseTable"/>
                    <w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="CFE2F3"/></w:tcPr></w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="ChildTable"/>
                        <w:tblLook w:firstRow="1" w:firstColumn="0" w:noHBand="1" w:noVBand="1"/>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc>
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

        DocxTableCell firstHeader = document.Tables[0].Rows[0].Cells[0];
        DocxTableCell firstBody = document.Tables[0].Rows[1].Cells[0];
        DocxParagraph headerParagraph = firstHeader.Paragraphs.Single();

        TestAssert.Equal("ChildTable", document.Tables[0].StyleId ?? string.Empty);
        TestAssert.Equal("CFE2F3", firstHeader.FillHex ?? string.Empty);
        TestAssert.Equal("center", firstHeader.VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("D9EAD3", firstBody.FillHex ?? string.Empty);
        TestAssert.Equal("240", firstHeader.Margins.LeftValue ?? string.Empty);
        TestAssert.Equal(DocxTextAlignment.Right, headerParagraph.Alignment);
        TestAssert.True(headerParagraph.Runs.Single().Bold, "Inherited base table run style should apply bold.");
        TestAssert.Equal("336699", headerParagraph.Runs.Single().ColorHex ?? string.Empty);
    }

    public static void DocxReaderTableStyleAppliesTableProperties()
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
                  <w:style w:type="table" w:styleId="IndentedTable">
                    <w:tblPr>
                      <w:tblLayout w:type="fixed"/>
                      <w:tblW w:w="2880" w:type="dxa"/>
                      <w:tblInd w:w="360" w:type="dxa"/>
                      <w:tblCellSpacing w:w="120" w:type="dxa"/>
                    </w:tblPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="IndentedTable"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>Inherited</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="IndentedTable"/>
                        <w:tblLayout w:type="autofit"/>
                        <w:tblInd w:w="720" w:type="dxa"/>
                        <w:tblCellSpacing w:w="240" w:type="dxa"/>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>Direct</w:t></w:r></w:p></w:tc></w:tr>
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
        TestAssert.Equal(144d, document.Tables[0].PreferredWidthPoints ?? 0d);
        TestAssert.Equal("360", document.Tables[0].IndentValue ?? string.Empty);
        TestAssert.Equal(18d, document.Tables[0].IndentPoints ?? 0d);
        TestAssert.Equal("120", document.Tables[0].CellSpacingValue ?? string.Empty);
        TestAssert.Equal(6d, document.Tables[0].CellSpacingPoints ?? 0d);
        TestAssert.Equal("autofit", document.Tables[1].LayoutValue ?? string.Empty);
        TestAssert.Equal("2880", document.Tables[1].PreferredWidthValue ?? string.Empty);
        TestAssert.Equal("720", document.Tables[1].IndentValue ?? string.Empty);
        TestAssert.Equal(36d, document.Tables[1].IndentPoints ?? 0d);
        TestAssert.Equal("240", document.Tables[1].CellSpacingValue ?? string.Empty);
        TestAssert.Equal(12d, document.Tables[1].CellSpacingPoints ?? 0d);
    }

    public static void DocxReaderTableStyleAppliesParagraphAndRunProperties()
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
                  <w:style w:type="table" w:styleId="TextTable">
                    <w:pPr><w:spacing w:after="0"/></w:pPr>
                    <w:rPr><w:sz w:val="22"/></w:rPr>
                    <w:tblStylePr w:type="firstCol">
                      <w:pPr><w:jc w:val="right"/></w:pPr>
                      <w:rPr><w:i/><w:caps/><w:color w:val="4472C4"/></w:rPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="TextTable"/>
                        <w:tblLook w:firstColumn="1" w:firstRow="0" w:noHBand="1" w:noVBand="1"/>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:tc>
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

        DocxParagraph firstParagraph = document.Tables[0].Rows[0].Cells[0].Paragraphs.Single();
        DocxParagraph secondParagraph = document.Tables[0].Rows[0].Cells[1].Paragraphs.Single();
        TestAssert.Equal(DocxTextAlignment.Right, firstParagraph.Alignment);
        TestAssert.Equal("right", firstParagraph.AlignmentValue ?? string.Empty);
        TestAssert.Equal(6d, firstParagraph.SpacingAfterPoints);
        TestAssert.Equal(11d, firstParagraph.Runs.Single().FontSize);
        TestAssert.True(firstParagraph.Runs.Single().Italic, "First-column table run style should apply italic.");
        TestAssert.True(firstParagraph.Runs.Single().AllCaps, "First-column table run style should apply all-caps.");
        TestAssert.Equal("FIRST", firstParagraph.Runs.Single().Text);
        TestAssert.Equal("4472C4", firstParagraph.Runs.Single().ColorHex ?? string.Empty);
        TestAssert.Equal(0d, secondParagraph.SpacingAfterPoints);
        TestAssert.Equal(11d, secondParagraph.Runs.Single().FontSize);
    }

    public static void DocxReaderTableStyleAppliesComplexScriptRunPropertiesByScriptSlot()
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
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblStylePr w:type="firstRow">
                      <w:rPr><w:bCs/><w:iCs/><w:rFonts w:cs="Complex Face"/></w:rPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="StyledTable"/></w:tblPr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A&#x05D0;B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        DocxDocument document = new DocxReader().Read(OoxPackage.Open(input));

        IReadOnlyList<DocxTextRun> runs = document.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single().Runs;
        TestAssert.Equal(3, runs.Count);
        TestAssert.Equal("A", runs[0].Text);
        TestAssert.True(!runs[0].Bold && !runs[0].Italic, "Latin text should not inherit complex-script bold/italic.");
        TestAssert.Equal("\u05D0", runs[1].Text);
        TestAssert.True(runs[1].Bold && runs[1].Italic, "Complex-script text should use bCs/iCs from the table style.");
        TestAssert.Equal("Complex Face", runs[1].FontFamily ?? string.Empty);
        TestAssert.Equal("B", runs[2].Text);
        TestAssert.True(!runs[2].Bold && !runs[2].Italic, "Latin text after the complex-script segment should return to Latin run properties.");
    }

    public static void DocxReaderTableStyleRunPropertiesStayBelowParagraphAndCharacterStyles()
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
                  <w:style w:type="paragraph" w:styleId="BodyAccent"><w:rPr><w:b/><w:color w:val="00AA00"/><w:sz w:val="26"/></w:rPr></w:style>
                  <w:style w:type="character" w:styleId="StrongRed"><w:rPr><w:color w:val="CC0000"/><w:sz w:val="28"/></w:rPr></w:style>
                  <w:style w:type="table" w:styleId="TextTable">
                    <w:rPr><w:sz w:val="22"/></w:rPr>
                    <w:tblStylePr w:type="firstCol"><w:rPr><w:i/><w:caps/><w:color w:val="4472C4"/></w:rPr></w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="TextTable"/><w:tblLook w:firstColumn="1" w:firstRow="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="BodyAccent"/></w:pPr><w:r><w:rPr><w:rStyle w:val="StrongRed"/></w:rPr><w:t>Conflict</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        DocxDocument document = new DocxReader().Read(OoxPackage.Open(input));
        DocxTextRun run = document.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single().Runs.Single();

        TestAssert.Equal(14d, run.FontSize);
        TestAssert.Equal("CC0000", run.ColorHex ?? string.Empty);
        TestAssert.True(run.Bold, "Paragraph-style run properties should remain above table-style run properties.");
        TestAssert.True(run.Italic, "Non-conflicting table-style italic should still apply.");
        TestAssert.True(run.AllCaps, "Non-conflicting table-style caps should still apply.");
        TestAssert.Equal("CONFLICT", run.Text);
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
                    <w:tblPr><w:tblStyleRowBandSize w:val="2"/><w:tblStyleColBandSize w:val="2"/></w:tblPr>
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
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Wide band</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Wide band</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Next band</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Next band</w:t></w:r></w:p></w:tc>
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
        TestAssert.Equal("00FF00", document.Tables[0].Rows[2].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("0000FF", document.Tables[0].Rows[3].Cells[0].FillHex ?? string.Empty);
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

    public static void DocxSyntheticTableCellBordersUseOfficeLikeFilledStrips()
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
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains(" re f", pdf);
        TestAssert.DoesNotContain("2 w", pdf);
        TestAssert.DoesNotContain(" l S", pdf);
        TestAssert.DoesNotContain(" re S", pdf);
    }

    public static void DocxSyntheticTableCollapsedBorderIntersectionsUseFilledNodes()
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
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideH w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideV w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
                      <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.48 0.48 re f", pdf);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 21, "Expected collapsed border strips plus grid-intersection nodes.");
    }

    public static void DocxSyntheticTableCellLogicalBordersRenderInLeftToRightLayout()
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
                          <w:tcPr><w:tcBorders><w:start w:val="single" w:color="0000FF" w:sz="8"/><w:end w:val="single" w:color="FF0000" w:sz="8"/></w:tcBorders></w:tcPr>
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
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 2, "Expected logical borders to render as filled rectangle strips.");
        TestAssert.DoesNotContain(" l S", pdf);
    }

    public static void DocxSyntheticTableSharedVerticalBorderRendersOnceAndHonorsNil()
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
                      <w:tblPr><w:tblBorders><w:insideV w:val="single" w:color="0000FF" w:sz="8"/></w:tblBorders></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p/></w:tc>
                        <w:tc><w:p/></w:tc>
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
        TestAssert.Equal(1, pdf.Split("0 0 1 rg", StringSplitOptions.None).Length - 1);

        string nilInput = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                      <w:tblPr><w:tblBorders><w:insideV w:val="single" w:color="0000FF" w:sz="8"/></w:tblBorders></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p/></w:tc>
                        <w:tc><w:tcPr><w:tcBorders><w:left w:val="nil"/></w:tcBorders></w:tcPr><w:p/></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string nilOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(nilInput, nilOutput);

        string nilPdf = File.ReadAllText(nilOutput, Encoding.ASCII);
        TestAssert.DoesNotContain("0 0 1 rg", nilPdf);
    }

    public static void DocxSyntheticTableSharedHorizontalBorderRendersOnceAndHonorsNil()
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
                      <w:tblPr><w:tblBorders><w:insideH w:val="single" w:color="0000FF" w:sz="8"/></w:tblBorders></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p/></w:tc></w:tr>
                      <w:tr><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, pdf.Split("0 0 1 rg", StringSplitOptions.None).Length - 1);

        string nilInput = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                      <w:tblPr><w:tblBorders><w:insideH w:val="single" w:color="0000FF" w:sz="8"/></w:tblBorders></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p/></w:tc></w:tr>
                      <w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="nil"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string nilOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(nilInput, nilOutput);

        string nilPdf = File.ReadAllText(nilOutput, Encoding.ASCII);
        TestAssert.DoesNotContain("0 0 1 rg", nilPdf);
    }

    public static void DocxSyntheticTableSharedHorizontalBorderUsesOverlappingCells()
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
                      <w:tblPr><w:tblBorders><w:insideH w:val="single" w:color="0000FF" w:sz="8"/></w:tblBorders></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr>
                      <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(2, pdf.Split("0 0 1 rg", StringSplitOptions.None).Length - 1);
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

    public static void DocxReaderInfersMissingTableGridFromLogicalGridSpans()
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
                      <w:tr>
                        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Span</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Tail</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxTable table = new DocxReader().Read(package).Tables[0];

        TestAssert.True(!table.HasExplicitGrid, "Missing tblGrid should remain distinguishable from authored grid geometry.");
        TestAssert.Equal(3, table.ColumnWidthsPoints.Count);
    }

    public static void DocxReaderTableCellPreservesVerticalMergeToken()
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
                      <w:tblGrid><w:gridCol w:w="720"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>Bottom</w:t></w:r></w:p></w:tc>
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

        DocxTableCell restart = document.Tables[0].Rows[0].Cells[0];
        DocxTableCell continuation = document.Tables[0].Rows[1].Cells[0];
        TestAssert.True(restart.HasVerticalMerge, "Reader should preserve vMerge restart presence.");
        TestAssert.Equal("restart", restart.VerticalMergeValue ?? string.Empty);
        TestAssert.True(continuation.HasVerticalMerge, "Reader should preserve val-less vMerge continuation presence.");
        TestAssert.True(continuation.VerticalMergeValue is null, "Val-less vMerge continuation should keep its missing value distinct from restart.");
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

    public static void DocxReaderTableLogicalBordersApplyOuterEdgesInLeftToRightLayout()
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
                          <w:start w:val="single" w:color="123456" w:sz="8"/>
                          <w:end w:val="single" w:color="654321" w:sz="8"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
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

        TestAssert.Equal("123456", borders.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("654321", borders.Single(border => border.Edge == "right").Color ?? string.Empty);
    }

    public static void DocxReaderDirectTableBordersOverrideTableStyleCellBordersPerEdge()
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
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblPr>
                      <w:tblBorders>
                        <w:left w:val="single" w:color="333333" w:sz="12"/>
                        <w:bottom w:val="single" w:color="222222" w:sz="10"/>
                      </w:tblBorders>
                    </w:tblPr>
                    <w:tblStylePr w:type="firstRow">
                      <w:tcPr>
                        <w:tcBorders>
                          <w:top w:val="single" w:color="FF0000" w:sz="18"/>
                          <w:right w:val="single" w:color="AA0000" w:sz="18"/>
                        </w:tcBorders>
                      </w:tcPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="StyledTable"/>
                        <w:tblBorders>
                          <w:top w:val="nil"/>
                          <w:right w:val="single" w:color="0000FF" w:sz="6"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
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
        TestAssert.Equal("nil", borders.Single(border => border.Edge == "top").Value ?? string.Empty);
        TestAssert.Equal("0000FF", borders.Single(border => border.Edge == "right").Color ?? string.Empty);
        TestAssert.Equal("333333", borders.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("222222", borders.Single(border => border.Edge == "bottom").Color ?? string.Empty);
    }

    public static void DocxReaderTableStyleAppliesConditionalCellVerticalAlignment()
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
                  <w:style w:type="table" w:styleId="AlignedTable">
                    <w:tcPr><w:vAlign w:val="center"/></w:tcPr>
                    <w:tblStylePr w:type="firstRow">
                      <w:tcPr><w:vAlign w:val="bottom"/></w:tcPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="AlignedTable"/>
                        <w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Inherited</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:vAlign w:val="top"/></w:tcPr><w:p><w:r><w:t>Direct</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Whole</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Whole2</w:t></w:r></w:p></w:tc>
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

        IReadOnlyList<DocxTableRow> rows = document.Tables[0].Rows;
        TestAssert.Equal("bottom", rows[0].Cells[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("top", rows[0].Cells[1].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("center", rows[1].Cells[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("center", rows[1].Cells[1].VerticalAlignmentValue ?? string.Empty);
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

    public static void DocxSyntheticTableCellPatternShadingUsesPdfTilingPattern()
    {
        string[] shadingValues =
        [
            "horzStripe",
            "thinHorzStripe",
            "vertStripe",
            "thinVertStripe",
            "diagStripe",
            "thinDiagStripe",
            "reverseDiagStripe",
            "thinReverseDiagStripe"
        ];
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
                      <w:tblGrid>
                        <w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/>
                        <w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/>
                      </w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:shd w:val="horzStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinHorzStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="vertStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinVertStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="diagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>5</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinDiagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>6</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="reverseDiagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>7</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinReverseDiagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>8</w:t></w:r></w:p></w:tc>
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
        TestAssert.Contains("/PatternType 1", pdf);
        TestAssert.Contains("/Pattern cs", pdf);
        TestAssert.Contains("/ImPattern", pdf);
        TestAssert.True(CountOccurrences(pdf, "/PatternType 1") >= shadingValues.Length, "Each supported DOCX stripe family should have a distinct tiling pattern resource.");
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

    public static void DocxReaderTableCellPreservesBodyFlowBreakElements()
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
                      <w:tr><w:tc>
                        <w:p><w:r><w:t>Left</w:t><w:br w:type="column"/><w:t>Right</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        var diagnostics = new List<OoxPdfDiagnostic>();

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package, diagnostics.Add);

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_MANUAL_BREAK"), "Visible table-cell column breaks should be preserved structurally without a stale unsupported diagnostic.");

        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];
        TestAssert.Equal("Left Right", cell.Text);
        TestAssert.Equal(3, cell.BodyElements.Count);
        TestAssert.Equal(2, cell.Paragraphs.Count);
        TestAssert.Equal("Left", cell.Paragraphs[0].Runs.Single().Text);
        TestAssert.True(cell.BodyElements[1] is DocxManualBreakElement, "The table-cell body stream should preserve the typed column break.");
        TestAssert.Equal("Right", cell.Paragraphs[1].Runs.Single().Text);

        DocxStructureTableCellSnapshot cellSnapshot = DocxStructureSnapshot.FromDocument(document).Tables.Single().Rows.Single().Cells.Single();
        TestAssert.Equal(3, cellSnapshot.BodyElementCount);
        TestAssert.Equal(1, cellSnapshot.ManualBreakElementCount);
        TestAssert.Equal(0, cellSnapshot.PageBreakElementCount);
    }

    public static void DocxReaderTableInventoryIncludesNestedTableCellBodies()
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
                      <w:tr><w:tc>
                        <w:p><w:r><w:t>Outer</w:t></w:r></w:p>
                        <w:tbl>
                          <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                          <w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr>
                        </w:tbl>
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

        TestAssert.Equal(1, document.BodyElements.OfType<DocxTableElement>().Count());
        TestAssert.Equal(2, document.Tables.Count);
        TestAssert.Equal("Outer", document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Runs.Single().Text);
        TestAssert.Equal("Nested", document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Runs.Single().Text);
        TestAssert.True(document.Tables[0].Rows[0].Cells[0].BodyElements[1] is DocxTableElement, "The top-level body stream should not flatten nested table structure.");
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
        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal("Item", line.Segments[2].Text);
        TestAssert.True(line.Segments[2].X > line.Segments[0].X, "Numbered table-cell text should be segmented after the list label.");
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
                      <w:tblPr>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
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
        int firstText = FirstPdfTextShowIndex(pdf);
        int tableGrid = pdf.IndexOf(" re f", StringComparison.Ordinal);
        int lastText = LastPdfTextShowIndex(pdf);
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
                      <w:tblPr>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
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
        TestAssert.Contains("72.48 683.52 143.52 0.48 re f", pdf);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxTableRow row = new DocxReader().Read(package).Tables[0].Rows[0];
        TestAssert.Equal("720", row.HeightValue ?? string.Empty);
        TestAssert.True(row.HeightRuleValue is null, "Missing hRule should stay distinct from exact/auto.");
    }

    public static void DocxReaderTableRowPreservesPropertyExceptionCellMargins()
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
                      <w:tr>
                        <w:tblPrEx>
                          <w:tblCellMar>
                            <w:top w:w="0" w:type="dxa"/>
                            <w:bottom w:w="0" w:type="dxa"/>
                          </w:tblCellMar>
                        </w:tblPrEx>
                        <w:tc><w:p/></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxTableRow row = new DocxReader().Read(package).Tables[0].Rows[0];

        TestAssert.True(row.TablePropertyExceptionCellMargins is not null, "Expected row-level tblPrEx cell margins to stay distinct from cell margins.");
        TestAssert.Equal("0", row.TablePropertyExceptionCellMargins!.TopValue ?? string.Empty);
        TestAssert.Equal("0", row.TablePropertyExceptionCellMargins.BottomValue ?? string.Empty);
    }

    public static void DocxReaderTableRowPreservesCantSplit()
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
                      <w:tr>
                        <w:trPr><w:cantSplit w:val="1"/></w:trPr>
                        <w:tc><w:p/></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxTableRow row = new DocxReader().Read(package).Tables[0].Rows[0];

        TestAssert.True(row.CantSplit, "Expected w:cantSplit to be preserved for row-fragment pagination.");
        TestAssert.Equal("1", row.CantSplitValue ?? string.Empty);
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

    public static void DocxTableLayoutStageManualColumnBreakAdvancesPageInSingleColumn()
    {
        DocxTable first = CreateSingleCellTable("first", rowHeight: 20d);
        DocxTable second = CreateSingleCellTable("second", rowHeight: 20d);
        DocxDocument document = CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxManualBreakElement("runBreak", "column"),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded: null);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageContinuesRowsInActiveColumnFrame()
    {
        var firstRow = new DocxTableRow([new DocxTableCell("first", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 100d);
        var secondRow = new DocxTableRow([new DocxTableCell("second", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 100d);
        var table = new DocxTable(null, [60d], [firstRow, secondRow]);
        DocxPageSettings sectionSettings = DocxPageSettings.Empty with
        {
            WidthValue = "4000",
            HeightValue = "4000",
            MarginLeftValue = "360",
            MarginRightValue = "360",
            MarginTopValue = "360",
            MarginBottomValue = "360"
        };
        var document = new DocxDocument(
            300d,
            300d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxTableElement(table),
                new DocxSectionBreakElement(sectionSettings, "nextPage", "2", "1", "360", [])
            ],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);

        TestAssert.Equal(1, layout.Pages.Count);
        TestAssert.Equal(2, rows.Length);
        TestAssert.Equal(18d, rows[0].Table.TableX);
        TestAssert.Equal(109d, rows[1].Table.TableX);
        TestAssert.Equal(0, snapshot.Pages[0].Items[0].ColumnIndex ?? -1);
        TestAssert.Equal(1, snapshot.Pages[0].Items[1].ColumnIndex ?? -1);
    }

    public static void DocxTableLayoutStageRunPageBreakParagraphConsumesLineBox()
    {
        DocxTable first = CreateSingleCellTable("first", rowHeight: 150d);
        DocxTable second = CreateSingleCellTable("second", rowHeight: 20d);
        DocxParagraph marker = CreateDocxLayoutParagraph("marker", 10d, 20d);
        DocxParagraph breakParagraph = CreateDocxLayoutParagraph("", 10d, 20d);
        DocxDocument document = CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxParagraphElement(marker),
            new DocxPageBreakElement("runBreak", "page", breakParagraph),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded: null);

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(0, layout.Pages[1].Items.Count);
        TestAssert.Equal(1, layout.Pages[2].Items.OfType<DocxTableRowLayout>().Count());
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

    public static void DocxTableLayoutStageSplitsTallRowsAcrossPagesByDefault()
    {
        DocxParagraph firstParagraph = CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] secondParagraphs = Enumerable.Range(1, 8)
            .Select(index => CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", secondParagraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [first, second]);
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

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(2, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, firstPageRows[1].RowIndex);
        TestAssert.Equal(0, firstPageRows[1].FragmentIndex);
        TestAssert.Equal(2, firstPageRows[1].FragmentCount);
        TestAssert.Equal("PageBoundary", firstPageRows[1].FragmentReason);
        TestAssert.Equal(80d, firstPageRows[1].FullRowHeight);
        TestAssert.Equal(0d, firstPageRows[1].FragmentOffsetFromRowTop);
        TestAssert.Equal(20d, firstPageRows[1].Height);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentIndex);
        TestAssert.Equal(2, secondPageRows[0].FragmentCount);
        TestAssert.Equal("PageBoundary", secondPageRows[0].FragmentReason);
        TestAssert.Equal(80d, secondPageRows[0].FullRowHeight);
        TestAssert.Equal(20d, secondPageRows[0].FragmentOffsetFromRowTop);
        TestAssert.Equal(60d, secondPageRows[0].Height);
        TestAssert.True(firstPageRows[1].Cells[0].TextLines.Count > 0, "The first split fragment should own its visible row text.");
        TestAssert.True(secondPageRows[0].Cells[0].TextLines.Count > 0, "The continuation split fragment should own its visible row text.");
        TestAssert.Equal(8, firstPageRows[1].Cells[0].TextLines.Count + secondPageRows[0].Cells[0].TextLines.Count);

        DocxTableSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout).Tables.Single();
        TestAssert.Equal(1, snapshot.FragmentedRowCount);
        TestAssert.Equal(2, snapshot.FragmentedRowLayoutCount);
        TestAssert.Equal(2, snapshot.MaxRowFragmentCount);
        DocxTableRowSnapshot[] splitRowSnapshots = DocxLayoutSnapshot.FromLayout(layout).Pages
            .SelectMany(page => page.TableRows)
            .Where(row => row.RowIndex == 1)
            .OrderBy(row => row.FragmentIndex)
            .ToArray();
        TestAssert.Equal(80d, splitRowSnapshots[0].FullRowHeight);
        TestAssert.Equal("PageBoundary", splitRowSnapshots[0].FragmentReason);
        TestAssert.Equal(0d, splitRowSnapshots[0].FragmentOffsetFromRowTop);
        TestAssert.Equal(80d, splitRowSnapshots[1].FullRowHeight);
        TestAssert.Equal("PageBoundary", splitRowSnapshots[1].FragmentReason);
        TestAssert.Equal(20d, splitRowSnapshots[1].FragmentOffsetFromRowTop);
    }

    public static void DocxTableLayoutStageRepeatsHeaderRowsBeforeSplitRowContinuations()
    {
        DocxParagraph headerParagraph = CreateDocxLayoutParagraph("Header", 10d, 10d);
        DocxParagraph fillerParagraph = CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var header = new DocxTableRow([new DocxTableCell("Header", [headerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 10d, IsHeader: true);
        var filler = new DocxTableRow([new DocxTableCell("Filler", [fillerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 50d);
        var split = new DocxTableRow([new DocxTableCell("Split", splitParagraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [header, filler, split]);
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

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(3, firstPageRows.Length);
        TestAssert.Equal(2, secondPageRows.Length);
        TestAssert.Equal(0, firstPageRows[0].RowIndex);
        TestAssert.True(firstPageRows[0].IsHeader, "The first page should keep the table header before body rows.");
        TestAssert.Equal(2, firstPageRows[2].RowIndex);
        TestAssert.Equal(0, firstPageRows[2].FragmentIndex);
        TestAssert.Equal(2, firstPageRows[2].FragmentCount);
        TestAssert.Equal(0, secondPageRows[0].RowIndex);
        TestAssert.True(secondPageRows[0].IsHeader, "Split-row continuations should repeat table headers before the carried fragment.");
        TestAssert.Equal(2, secondPageRows[1].RowIndex);
        TestAssert.Equal(1, secondPageRows[1].FragmentIndex);
        TestAssert.Equal(2, secondPageRows[1].FragmentCount);
        TestAssert.Equal(60d, secondPageRows[1].Height);
    }

    public static void DocxTableLayoutStageKeepsInlineImagesInSplitRowFragments()
    {
        var image = new DocxInlineImage(20d, 20d, "image/png", [0x89, 0x50, 0x4E, 0x47], "word/media/image1.png");
        DocxParagraph fillerParagraph = CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 7)
            .Select(index => CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .Append(new DocxParagraph(
                [],
                [image],
                null,
                DocxTextAlignment.Left,
                null,
                0d,
                0d,
                1d,
                10d,
                DocxParagraphSpacing.Empty,
                DocxParagraphKeepRules.Empty,
                null))
            .ToArray();
        var filler = new DocxTableRow([new DocxTableCell("Filler", [fillerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var split = new DocxTableRow([new DocxTableCell("Split", splitParagraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], 100d);
        var table = new DocxTable(null, [60d], [filler, split]);
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

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] splitFragments = layout.Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .Where(row => row.RowIndex == 1)
            .ToArray();

        TestAssert.Equal(2, splitFragments.Length);
        TestAssert.True(splitFragments.All(fragment => fragment.FragmentCount == 2), "The image-owning body row should split into two fragments.");
        TestAssert.True(splitFragments.Sum(fragment => fragment.Cells.Sum(cell => cell.InlineImages.Count)) > 0, "Split row fragments should keep overlapping inline image layouts for the renderer clip path.");
    }

    public static void DocxTableLayoutStageClipsVerticalMergeRestartToSplitRowFragments()
    {
        DocxParagraph fillerParagraph = CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var filler = new DocxTableRow([new DocxTableCell("Filler", [fillerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var restart = new DocxTableRow([
            new DocxTableCell(
                "Merged",
                splitParagraphs,
                null,
                null,
                null,
                null,
                [],
                DocxTableCellMargins.Empty,
                HasVerticalMerge: true,
                VerticalMergeValue: "restart")
        ], 100d);
        var continuation = new DocxTableRow([
            new DocxTableCell(
                "Continuation",
                [],
                null,
                null,
                null,
                null,
                [],
                DocxTableCellMargins.Empty,
                HasVerticalMerge: true)
        ], 20d);
        var table = new DocxTable(null, [60d], [filler, restart, continuation]);
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

        DocxTableRowLayout[] splitFragments = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .Where(row => row.RowIndex == 1)
            .ToArray();

        TestAssert.Equal(2, splitFragments.Length);
        TestAssert.True(splitFragments.All(fragment => fragment.FragmentCount == 2), "The merged restart row should still split into physical row fragments.");
        TestAssert.True(splitFragments.All(fragment => fragment.Cells[0].Y == fragment.Y), "A split merged restart cell should use the fragment top as its visible clip.");
        TestAssert.True(splitFragments.All(fragment => fragment.Cells[0].Height == fragment.Height), "A split merged restart cell should clip to the fragment height, not the full cross-row merged span.");
        TestAssert.True(splitFragments.All(fragment => fragment.Cells[0].TextLines.Count != 0), "Full merged-cell text coordinates should remain available for fragment clipping.");
    }

    public static void DocxTableLayoutStageNormalizesPlainCellTextThroughSharedParagraphs()
    {
        var table = new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell("Plain", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cell = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        DocxTextLineLayout line = cell.TextLines.Single();
        TestAssert.Equal("Plain", line.Text);
        TestAssert.Equal(11d, line.FontSize);
    }

    public static void DocxTableLayoutStageHonorsCantSplitRowsAtPageBoundary()
    {
        var first = new DocxTableRow([new DocxTableCell("First", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 80d, CantSplit: true, CantSplitValue: "1");
        var table = new DocxTable(null, [60d], [first, second]);
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
        DocxTableRowLayout firstPageRow = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        DocxTableRowLayout secondPageRow = layout.Pages[1].Items.OfType<DocxTableRowLayout>().Single();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(0, firstPageRow.RowIndex);
        TestAssert.Equal(1, secondPageRow.RowIndex);
        TestAssert.Equal(0, secondPageRow.FragmentIndex);
        TestAssert.Equal(1, secondPageRow.FragmentCount);
        TestAssert.True(secondPageRow.CantSplit, "w:cantSplit rows should move whole instead of creating fragments.");
    }

    public static void DocxTableLayoutStageKeepsFollowingParagraphAdjacent()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("After", 10d, null, false, false, false, null, null)],
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
            null);
        var table = new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell("Cell", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table), new DocxParagraphElement(paragraph)], [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        DocxTableRowLayout row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        DocxTextLineLayout following = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single();
        double expectedBaselineY = row.Y - DocxLineMetrics.ResolveBodyBaselineOffset(10d, 10d, hasExplicitLineSpacing: true);
        TestAssert.Equal(Math.Round(expectedBaselineY, 3), Math.Round(following.BaselineY, 3));
    }

    public static void DocxTableLayoutStageKeepsParagraphWithFollowingTableFirstRow()
    {
        DocxParagraph filler = CreateDocxLayoutParagraph("Filler", 10d, 60d);
        DocxParagraph heading = CreateDocxLayoutParagraph(
            "Heading",
            10d,
            20d,
            new DocxParagraphKeepRules(true, "1", null, null, null, null));
        var table = new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell("Cell", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
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
            [new DocxParagraphElement(filler), new DocxParagraphElement(heading), new DocxTableElement(table)],
            [filler, heading],
            [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("Filler", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal("Heading", layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageAppliesVerticalMergeGeometry()
    {
        var restart = new DocxTableCell(
            "Merged",
            [],
            null,
            null,
            null,
            null,
            [new DocxTableCellBorder("bottom", "single", "000000", "8")],
            DocxTableCellMargins.Empty,
            HasVerticalMerge: true,
            VerticalMergeValue: "restart");
        var continuation = new DocxTableCell(
            "Continuation",
            [],
            null,
            null,
            null,
            null,
            [new DocxTableCellBorder("top", "single", "000000", "8")],
            DocxTableCellMargins.Empty,
            HasVerticalMerge: true);
        var table = new DocxTable(
            null,
            [60d],
            [
                new DocxTableRow([restart], 20d),
                new DocxTableRow([continuation], 30d)
            ]);
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

        DocxTableRowLayout[] rows = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .ToArray();

        TestAssert.Equal(2, rows.Length);
        TestAssert.True(Math.Abs(rows[0].Cells[0].Y - 38.08d) < 0.001d, $"Expected merged restart y near 38.08pt, got {rows[0].Cells[0].Y.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(rows[0].Cells[0].Height - 51.92d) < 0.001d, $"Expected merged restart height near 51.92pt, got {rows[0].Cells[0].Height.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.Equal(DocxTableCellVisualOwnership.OwnCell, rows[0].Cells[0].VisualOwnership);
        TestAssert.True(ReferenceEquals(restart, rows[0].Cells[0].VisualCell), "Restart cells should keep themselves as visual cells.");
        TestAssert.True(rows[1].Cells[0].IsVerticalMergeContinuation, "Continuation cell should be layout-visible but skipped by rendering.");
        TestAssert.Equal(DocxTableCellVisualOwnership.VerticalMergeOwner, rows[1].Cells[0].VisualOwnership);
        TestAssert.Equal(0, rows[1].Cells[0].VerticalMergeOwner?.RowIndex ?? -1);
        TestAssert.Equal(0, rows[1].Cells[0].VerticalMergeOwner?.GridColumnIndex ?? -1);
        TestAssert.True(ReferenceEquals(restart, rows[1].Cells[0].VisualCell), "Continuation cells should expose the restart cell as their visual source.");
        TestAssert.True(Math.Abs(rows[1].Cells[0].Y - 38.08d) < 0.001d, $"Expected merge continuation y near 38.08pt, got {rows[1].Cells[0].Y.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(rows[1].Cells[0].Height - 30.96d) < 0.001d, $"Expected merge continuation height near 30.96pt, got {rows[1].Cells[0].Height.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void DocxTableLayoutStageCarriesVerticalMergeOwnerAcrossPages()
    {
        var filler = new DocxTableCell("Filler", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var restart = new DocxTableCell(
            "Merged",
            [CreateDocxLayoutParagraph("Merged", 10d, 10d)],
            "D9EAD3",
            "clear",
            "auto",
            null,
            [new DocxTableCellBorder("left", "single", "000000", "8")],
            DocxTableCellMargins.Empty,
            HasVerticalMerge: true,
            VerticalMergeValue: "restart");
        var continuation = new DocxTableCell(
            "Continuation",
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty,
            HasVerticalMerge: true);
        var table = new DocxTable(
            null,
            [60d],
            [
                new DocxTableRow([filler], 60d),
                new DocxTableRow([restart], 20d),
                new DocxTableRow([continuation], 30d)
            ]);
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

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout restartRow = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single(row => row.RowIndex == 1);
        DocxTableRowLayout continuationRow = layout.Pages[1].Items.OfType<DocxTableRowLayout>().Single(row => row.RowIndex == 2);
        DocxTableCellLayout continuationCell = continuationRow.Cells.Single();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.True(restartRow.Cells[0].Height > restartRow.Height, "The restart row should own the full merged span, even when it crosses the page boundary.");
        TestAssert.Equal(DocxTableCellVisualOwnership.OwnCell, restartRow.Cells[0].VisualOwnership);
        TestAssert.True(continuationCell.IsVerticalMergeContinuation, "The second-page row should remain marked as a merge continuation.");
        TestAssert.Equal(DocxTableCellVisualOwnership.VerticalMergeOwner, continuationCell.VisualOwnership);
        TestAssert.Equal(1, continuationCell.VerticalMergeOwner?.RowIndex ?? -1);
        TestAssert.Equal(0, continuationCell.VerticalMergeOwner?.GridColumnIndex ?? -1);
        TestAssert.True(ReferenceEquals(restart, continuationCell.VerticalMergeOwnerCell), "Continuation fragments should retain the restart cell as visual owner across pages.");
        TestAssert.True(ReferenceEquals(restart, continuationCell.VisualCell), "Continuation fragments should resolve their visual cell through layout ownership.");
        TestAssert.Equal("D9EAD3", continuationCell.VerticalMergeOwnerCell?.FillHex ?? string.Empty);
        TestAssert.Equal(60d, continuationCell.Y);
        TestAssert.Equal(30d, continuationCell.Height);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        DocxTableSnapshot tableSnapshot = snapshot.Tables.Single();
        TestAssert.True(tableSnapshot.HasVerticalMerge, "Table-level snapshots should expose vertical-merge presence.");
        TestAssert.Equal(2, tableSnapshot.AuthoredVerticalMergeCellCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeRestartCellCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeContinuationCellCount);
        TestAssert.Equal(1, tableSnapshot.LaidOutVerticalMergeContinuationCellCount);
        TestAssert.Equal(0, tableSnapshot.MissingVerticalMergeOwnerCellCount);

        DocxTableCellSnapshot continuationSnapshot = snapshot.Pages[1].TableRows.Single(row => row.RowIndex == 2).Cells.Single();
        TestAssert.Equal("VerticalMergeOwner", continuationSnapshot.VisualOwnership);
        TestAssert.Equal(1, continuationSnapshot.VerticalMergeOwnerRowIndex ?? -1);
        TestAssert.Equal(0, continuationSnapshot.VerticalMergeOwnerGridColumnIndex ?? -1);
        TestAssert.Equal(6, continuationSnapshot.TextLength);
        TestAssert.Equal(6, continuationSnapshot.VisualTextLength);
        TestAssert.Equal(1, continuationSnapshot.VisualParagraphCount);
        TestAssert.True(continuationCell.TextLines.Count != 0, "A page-crossing merge continuation should carry owner text lines for page-local clipped PDF emission.");
    }

    public static void DocxTableLayoutStageCarriesVerticalMergeOwnerAfterRepeatedHeader()
    {
        var header = new DocxTableRow(
            [new DocxTableCell("Header", [CreateDocxLayoutParagraph("Header", 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty)],
            10d,
            IsHeader: true);
        var filler = new DocxTableRow(
            [new DocxTableCell("Filler", [CreateDocxLayoutParagraph("Filler", 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty)],
            50d);
        var restart = new DocxTableCell(
            "Merged",
            [CreateDocxLayoutParagraph("Merged", 10d, 10d)],
            "D9EAD3",
            "clear",
            "auto",
            null,
            [new DocxTableCellBorder("left", "single", "000000", "8")],
            DocxTableCellMargins.Empty,
            HasVerticalMerge: true,
            VerticalMergeValue: "restart");
        var continuation = new DocxTableCell(
            "Continuation",
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty,
            HasVerticalMerge: true);
        var table = new DocxTable(
            null,
            [60d],
            [
                header,
                filler,
                new DocxTableRow([restart], 20d),
                new DocxTableRow([continuation], 30d)
            ]);
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

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableCellLayout continuationCell = secondPageRows.Single(row => row.RowIndex == 3).Cells.Single();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(2, secondPageRows.Length);
        TestAssert.Equal(0, secondPageRows[0].RowIndex);
        TestAssert.True(secondPageRows[0].IsHeader, "The second page should repeat the table header before the merge continuation.");
        TestAssert.Equal(3, secondPageRows[1].RowIndex);
        TestAssert.True(continuationCell.IsVerticalMergeContinuation, "The row after the repeated header should remain a merge continuation.");
        TestAssert.Equal(DocxTableCellVisualOwnership.VerticalMergeOwner, continuationCell.VisualOwnership);
        TestAssert.Equal(2, continuationCell.VerticalMergeOwner?.RowIndex ?? -1);
        TestAssert.Equal(0, continuationCell.VerticalMergeOwner?.GridColumnIndex ?? -1);
        TestAssert.True(ReferenceEquals(restart, continuationCell.VisualCell), "The repeated header must not become the visual owner of the merge continuation.");
        TestAssert.Equal("D9EAD3", continuationCell.VisualCell.FillHex ?? string.Empty);
        TestAssert.Equal(50d, continuationCell.Y);
        TestAssert.Equal(30d, continuationCell.Height);

        DocxTableCellSnapshot continuationSnapshot = DocxLayoutSnapshot.FromLayout(layout)
            .Pages[1]
            .TableRows
            .Single(row => row.RowIndex == 3)
            .Cells
            .Single();
        TestAssert.Equal("VerticalMergeOwner", continuationSnapshot.VisualOwnership);
        TestAssert.Equal(2, continuationSnapshot.VerticalMergeOwnerRowIndex ?? -1);
        TestAssert.Equal(6, continuationSnapshot.TextLength);
        TestAssert.Equal(6, continuationSnapshot.VisualTextLength);
        TestAssert.True(continuationCell.TextLines.Count != 0, "A merge continuation after a repeated header should carry owner text lines for page-local clipped PDF emission.");
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

    public static void DocxTableLayoutStageDoesNotClampExplicitDxaWidthToBody()
    {
        var table = new DocxTable(
            null,
            [110d, 110d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            PreferredWidthPoints: 220d,
            PreferredWidthValue: "4400",
            PreferredWidthType: "dxa");
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(110d, row.Cells[0].Width);
        TestAssert.Equal(120d, row.Cells[1].X);
        TestAssert.Equal(110d, row.Cells[1].Width);
        TestAssert.Equal(220d, row.Table.ResolvedTableWidth);
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

    public static void DocxTableLayoutStageDistributesMissingGridAcrossAvailableWidth()
    {
        var table = new DocxTable(
            null,
            [72d, 72d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            HasExplicitGrid: false);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(!row.Table.HasExplicitGrid, "Layout context should preserve missing-grid provenance after width resolution.");
        TestAssert.Equal(90d, row.Cells[0].Width);
        TestAssert.Equal(100d, row.Cells[1].X);
        TestAssert.Equal(90d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageDistributesMissingGridWithSpansAcrossLogicalColumns()
    {
        var table = new DocxTable(
            null,
            [72d, 72d, 72d],
            [new DocxTableRow([
                new DocxTableCell("wide", [], null, null, null, null, [], DocxTableCellMargins.Empty, GridSpan: 2, GridSpanValue: "2"),
                new DocxTableCell("tail", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            HasExplicitGrid: false);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(120d, row.Cells[0].Width);
        TestAssert.Equal(130d, row.Cells[1].X);
        TestAssert.Equal(60d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageInfersEmptyMissingGridFromRowSpans()
    {
        var table = new DocxTable(
            null,
            [],
            [new DocxTableRow([
                new DocxTableCell("wide", [], null, null, null, null, [], DocxTableCellMargins.Empty, GridSpan: 2, GridSpanValue: "2"),
                new DocxTableCell("tail", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)],
            HasExplicitGrid: false);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(120d, row.Cells[0].Width);
        TestAssert.Equal(130d, row.Cells[1].X);
        TestAssert.Equal(60d, row.Cells[1].Width);
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

    public static void DocxTableLayoutStageUsesCellBodyElementsForInlineImages()
    {
        var image = new DocxInlineImage(18d, 12d, "image/png", [1, 2, 3], "/word/media/image1.png");
        var paragraph = new DocxParagraph(
            [],
            [image],
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
        var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxParagraphElement(paragraph)]
        };
        var table = new DocxTable(null, [60d], [new DocxTableRow([cell], 36d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine()
            .Create(document, embedded: null)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(1, cellLayout.InlineImages.Count);
        TestAssert.Equal(18d, cellLayout.InlineImages[0].Width);
        TestAssert.Equal(12d, cellLayout.InlineImages[0].Height);
    }

    public static void DocxTableLayoutStageLaysOutNestedTableCellBodies()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        DocxParagraph nestedParagraph = CreateDocxLayoutParagraph("Nested", 10d, 10d);
        var nestedCell = new DocxTableCell("Nested", [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var nestedTable = new DocxTable(null, [50d], [new DocxTableRow([nestedCell], null)]);
        var outerCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var outerTable = new DocxTable(null, [80d], [new DocxTableRow([outerCell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(outerTable)], [outerTable, nestedTable]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font.Value.Font, "Nested".EnumerateRunes().Select(rune => rune.Value));

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded);
        DocxTableRowLayout outerRow = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        DocxTableCellLayout outerCellLayout = outerRow.Cells.Single();
        DocxTableRowLayout nestedRow = outerCellLayout.NestedRows.Single();
        DocxTextLineLayout nestedLine = nestedRow.Cells.Single().TextLines.Single();
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        var renderer = new DocxRenderer(new SingleResolutionFontResolver(font.Value.Resolution));
        DocxTextEmissionLineSnapshot[] emissionLines = renderer.InspectTextEmission(document).Lines
            .Where(line => !line.IsStaticStory)
            .ToArray();

        TestAssert.Equal("Nested", nestedLine.Text);
        TestAssert.True(outerRow.Height >= nestedRow.Height, "Outer table row height should include nested table content height.");
        TestAssert.True(nestedRow.Y >= outerCellLayout.Y && nestedRow.Y + nestedRow.Height <= outerCellLayout.Y + outerCellLayout.Height + 0.001d, "Nested table row should be placed inside the parent cell fragment.");
        TestAssert.Equal("Nested".Length, snapshot.Pages.Single().TableRows.Single().TextLength);
        TestAssert.True(emissionLines.Single().TextLength >= "Nested".Length, "Nested table-cell text should participate in renderer text-emission enumeration.");
    }

    public static void DocxTableLayoutStagePreservesMixedCellBodyOrderAroundNestedTables()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        DocxParagraph before = CreateDocxLayoutParagraph("Before", 10d, 10d);
        DocxParagraph nestedParagraph = CreateDocxLayoutParagraph("Nested", 10d, 10d);
        DocxParagraph after = CreateDocxLayoutParagraph("After", 10d, 10d);
        var nestedCell = new DocxTableCell("Nested", [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var nestedTable = new DocxTable(null, [50d], [new DocxTableRow([nestedCell], null)]);
        var outerCell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxParagraphElement(before), new DocxTableElement(nestedTable), new DocxParagraphElement(after)]
        };
        var outerTable = new DocxTable(null, [90d], [new DocxTableRow([outerCell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(outerTable)], [outerTable, nestedTable]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font.Value.Font, "BeforeNestedAfter".EnumerateRunes().Select(rune => rune.Value));

        DocxTableCellLayout outerCellLayout = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();
        DocxTextLineLayout[] outerLines = outerCellLayout.TextLines.ToArray();
        DocxTableRowLayout nestedRow = outerCellLayout.NestedRows.Single();

        TestAssert.Equal(2, outerLines.Length);
        TestAssert.Equal("Before", outerLines[0].Text);
        TestAssert.Equal("After", outerLines[1].Text);
        TestAssert.True(
            outerLines[0].BaselineY >= nestedRow.Y + nestedRow.Height,
            $"Paragraph before the nested table should remain above the nested table block. BeforeBaseline={outerLines[0].BaselineY}, NestedTop={nestedRow.Y + nestedRow.Height}, NestedBottom={nestedRow.Y}.");
        TestAssert.True(
            outerLines[1].BaselineY < nestedRow.Y,
            $"Paragraph after the nested table should be laid out below the nested table block. AfterBaseline={outerLines[1].BaselineY}, NestedTop={nestedRow.Y + nestedRow.Height}, NestedBottom={nestedRow.Y}.");
    }

    public static void DocxTableLayoutStageSplitsRowsAtCellPageBreakElements()
    {
        DocxParagraph before = CreateDocxLayoutParagraph("Before", 10d, 10d);
        DocxParagraph after = CreateDocxLayoutParagraph("After", 10d, 10d);
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement("runBreak", "page"),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] rowFragments = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(2, rowFragments.Length);
        TestAssert.Equal(0, rowFragments[0].FragmentIndex);
        TestAssert.Equal(1, rowFragments[1].FragmentIndex);
        TestAssert.Equal(2, rowFragments[0].FragmentCount);
        TestAssert.Equal(2, rowFragments[1].FragmentCount);
        TestAssert.Equal("CellPageBreak", rowFragments[0].FragmentReason);
        TestAssert.Equal("CellPageBreak", rowFragments[1].FragmentReason);
        TestAssert.Equal("Before", rowFragments[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", rowFragments[1].Cells.Single().TextLines.Single().Text);
        DocxTableRowSnapshot[] rowSnapshots = DocxLayoutSnapshot.FromLayout(layout).Pages
            .SelectMany(page => page.TableRows)
            .OrderBy(row => row.FragmentIndex)
            .ToArray();
        TestAssert.Equal("CellPageBreak", rowSnapshots[0].FragmentReason);
        TestAssert.Equal("CellPageBreak", rowSnapshots[1].FragmentReason);
    }

    public static void DocxTableLayoutStageMovesCellPageBreakSplitToNextPageWhenFirstFragmentDoesNotFit()
    {
        DocxParagraph before = CreateDocxLayoutParagraph("Before", 10d, 20d);
        DocxParagraph after = CreateDocxLayoutParagraph("After", 10d, 10d);
        var fillerCell = new DocxTableCell("Filler", [CreateDocxLayoutParagraph("Filler", 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty);
        var splitCell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement("runBreak", "page"),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([fillerCell], 70d), new DocxTableRow([splitCell], null)]);
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

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] rowFragments = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] thirdPageRows = layout.Pages[2].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(3, rowFragments.Length);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, thirdPageRows.Length);
        TestAssert.Equal(0, firstPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(0, secondPageRows[0].FragmentIndex);
        TestAssert.Equal("CellPageBreak", secondPageRows[0].FragmentReason);
        TestAssert.True(secondPageRows[0].Y - secondPageRows[0].Height >= 10d, "The first cell-page-break fragment must fit inside the new page content frame.");
        TestAssert.Equal(1, thirdPageRows[0].RowIndex);
        TestAssert.Equal(1, thirdPageRows[0].FragmentIndex);
        TestAssert.Equal("CellPageBreak", thirdPageRows[0].FragmentReason);
    }

    public static void DocxTableLayoutStageSplitsOversizedCellPageBreakFragmentAcrossFreshPages()
    {
        DocxParagraph[] beforeParagraphs = Enumerable.Range(1, 10)
            .Select(index => CreateDocxLayoutParagraph("Before" + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        DocxParagraph after = CreateDocxLayoutParagraph("After", 10d, 10d);
        var splitCell = new DocxTableCell(string.Empty, beforeParagraphs.Append(after).ToArray(), null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = beforeParagraphs
                .Select<DocxParagraph, DocxBodyElement>(paragraph => new DocxParagraphElement(paragraph))
                .Append(new DocxPageBreakElement("runBreak", "page"))
                .Append(new DocxParagraphElement(after))
                .ToArray()
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([splitCell], null)]);
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

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxTableRowLayout[] rowFragments = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(3, rowFragments.Length);
        for (int fragmentIndex = 0; fragmentIndex < rowFragments.Length; fragmentIndex++)
        {
            TestAssert.Equal(fragmentIndex, rowFragments[fragmentIndex].FragmentIndex);
            TestAssert.Equal(3, rowFragments[fragmentIndex].FragmentCount);
            TestAssert.Equal("CellPageBreak", rowFragments[fragmentIndex].FragmentReason);
        }

        TestAssert.Equal(80d, rowFragments[0].Height);
        TestAssert.Equal(20d, rowFragments[1].Height);
        TestAssert.Equal(10d, rowFragments[2].Height);
        TestAssert.Equal(8, rowFragments[0].Cells.Single().TextLines.Count);
        TestAssert.Equal("Before1|Before2|Before3|Before4|Before5|Before6|Before7|Before8", string.Join("|", rowFragments[0].Cells.Single().TextLines.Select(line => line.Text)));
        TestAssert.Equal("Before9|Before10", string.Join("|", rowFragments[1].Cells.Single().TextLines.Select(line => line.Text)));
        TestAssert.Equal("After", rowFragments[2].Cells.Single().TextLines.Single().Text);

        DocxTableRowSnapshot[] rowSnapshots = DocxLayoutSnapshot.FromLayout(layout).Pages
            .SelectMany(page => page.TableRows)
            .OrderBy(row => row.FragmentIndex)
            .ToArray();
        TestAssert.Equal(0d, rowSnapshots[0].FragmentOffsetFromRowTop);
        TestAssert.Equal(80d, rowSnapshots[1].FragmentOffsetFromRowTop);
        TestAssert.Equal(100d, rowSnapshots[2].FragmentOffsetFromRowTop);
        TestAssert.Equal(110d, rowSnapshots[0].FullRowHeight);
        TestAssert.Equal(110d, rowSnapshots[1].FullRowHeight);
        TestAssert.Equal(110d, rowSnapshots[2].FullRowHeight);
    }

    public static void DocxTableLayoutStageUsesEarliestCellPageBreakAsRowBoundary()
    {
        DocxParagraph earlyBefore = CreateDocxLayoutParagraph("EarlyBefore", 10d, 10d);
        DocxParagraph earlyAfter = CreateDocxLayoutParagraph("EarlyAfter", 10d, 10d);
        DocxParagraph laterFirst = CreateDocxLayoutParagraph("LaterFirst", 10d, 10d);
        DocxParagraph laterMiddle = CreateDocxLayoutParagraph("LaterMiddle", 10d, 10d);
        DocxParagraph laterAfter = CreateDocxLayoutParagraph("LaterAfter", 10d, 10d);
        var earlyCell = new DocxTableCell(string.Empty, [earlyBefore, earlyAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(earlyBefore),
                new DocxPageBreakElement("earlyBreak", "page"),
                new DocxParagraphElement(earlyAfter)
            ]
        };
        var laterCell = new DocxTableCell(string.Empty, [laterFirst, laterMiddle, laterAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(laterFirst),
                new DocxParagraphElement(laterMiddle),
                new DocxPageBreakElement("laterBreak", "page"),
                new DocxParagraphElement(laterAfter)
            ]
        };
        DocxTable table = new(null, [60d, 60d], [new DocxTableRow([earlyCell, laterCell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        TestAssert.Equal(3, rowFragments.Length);
        TestAssert.Equal("EarlyBefore", rowFragments[0].Cells[0].TextLines.Single().Text);
        TestAssert.Equal("EarlyAfter", rowFragments[1].Cells[0].TextLines.Single().Text);
        TestAssert.Equal("LaterFirst", rowFragments[0].Cells[1].TextLines.Single().Text);
        TestAssert.Equal("LaterMiddle", rowFragments[1].Cells[1].TextLines.Single().Text);
        TestAssert.Equal(0, rowFragments[2].Cells[0].TextLines.Count);
        TestAssert.Equal("LaterAfter", rowFragments[2].Cells[1].TextLines.Single().Text);
    }

    public static void DocxTableLayoutStageKeepsCellImagesOnAuthoredSideOfPageBreak()
    {
        var beforeImage = new DocxInlineImage(12d, 8d, "image/png", [1, 2, 3], "/word/media/before.png");
        var afterImage = new DocxInlineImage(10d, 6d, "image/png", [4, 5, 6], "/word/media/after.png");
        DocxParagraph before = CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { Images = [beforeImage] };
        DocxParagraph after = CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { Images = [afterImage] };
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement("runBreak", "page"),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        TestAssert.Equal(2, rowFragments.Length);
        DocxInlineImageLayout firstImage = rowFragments[0].Cells.Single().InlineImages.Single();
        DocxInlineImageLayout secondImage = rowFragments[1].Cells.Single().InlineImages.Single();
        TestAssert.Equal(0, firstImage.SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, secondImage.SourceParagraphIndex ?? -1);
        TestAssert.True(firstImage.Image == beforeImage, "The first row fragment should keep the image before the authored page break.");
        TestAssert.True(secondImage.Image == afterImage, "The second row fragment should keep the image after the authored page break.");
    }

    public static void DocxTableLayoutStageKeepsNestedTablesOnAuthoredSideOfPageBreak()
    {
        DocxTable beforeNestedTable = CreateSingleCellTable("Before", 12d);
        DocxTable afterNestedTable = CreateSingleCellTable("After", 12d);
        var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxTableElement(beforeNestedTable),
                new DocxPageBreakElement("runBreak", "page"),
                new DocxTableElement(afterNestedTable)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table, beforeNestedTable, afterNestedTable]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        TestAssert.Equal(2, rowFragments.Length);
        DocxTableRowLayout firstNestedRow = rowFragments[0].Cells.Single().NestedRows.Single();
        DocxTableRowLayout secondNestedRow = rowFragments[1].Cells.Single().NestedRows.Single();
        TestAssert.Equal(0, firstNestedRow.Table.TableIndex);
        TestAssert.Equal(1, secondNestedRow.Table.TableIndex);
        TestAssert.Equal("Before", firstNestedRow.Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", secondNestedRow.Cells.Single().TextLines.Single().Text);
    }

    public static void DocxTableLayoutStageKeepsTableCellColumnBreakInline()
    {
        DocxParagraph before = CreateDocxLayoutParagraph("Left", 10d, 10d);
        DocxParagraph after = CreateDocxLayoutParagraph("Right", 10d, 10d);
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxManualBreakElement("runBreak", "column"),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(1, cellLayout.TextLines.Count);
        TestAssert.Equal("LeftRight", cellLayout.TextLines.Single().Text);
    }

    public static void DocxTableLayoutStagePartitionsNestedTablesAcrossCompetingCellPageBreaks()
    {
        DocxParagraph earlyBefore = CreateDocxLayoutParagraph("EarlyBefore", 8d, 12d);
        DocxParagraph earlyAfter = CreateDocxLayoutParagraph("EarlyAfter", 8d, 12d);
        DocxTable beforeNestedTable = CreateSingleCellTable("Before", 12d);
        DocxTable middleNestedTable = CreateSingleCellTable("Middle", 12d);
        DocxTable afterNestedTable = CreateSingleCellTable("After", 10d);
        var earlyCell = new DocxTableCell(string.Empty, [earlyBefore, earlyAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(earlyBefore),
                new DocxPageBreakElement("earlyBreak", "page"),
                new DocxParagraphElement(earlyAfter)
            ]
        };
        var nestedCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxTableElement(beforeNestedTable),
                new DocxTableElement(middleNestedTable),
                new DocxPageBreakElement("laterBreak", "page"),
                new DocxTableElement(afterNestedTable)
            ]
        };
        DocxTable table = new(null, [60d, 90d], [new DocxTableRow([earlyCell, nestedCell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table, beforeNestedTable, middleNestedTable, afterNestedTable]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        TestAssert.Equal(3, rowFragments.Length);
        TestAssert.Equal("Before", rowFragments[0].Cells[1].NestedRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Middle", rowFragments[1].Cells[1].NestedRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", rowFragments[2].Cells[1].NestedRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal(0, rowFragments[0].Cells[1].NestedRows.Single().Table.TableIndex);
        TestAssert.Equal(1, rowFragments[1].Cells[1].NestedRows.Single().Table.TableIndex);
        TestAssert.Equal(2, rowFragments[2].Cells[1].NestedRows.Single().Table.TableIndex);
    }

    public static void DocxTableLayoutStageSplitsNestedTableRowsAcrossCompetingCellPageBreak()
    {
        DocxParagraph earlyBefore = CreateDocxLayoutParagraph("EarlyBefore", 10d, 20d);
        DocxParagraph earlyAfter = CreateDocxLayoutParagraph("EarlyAfter", 10d, 10d);
        DocxTableRow[] nestedRows = Enumerable.Range(1, 6)
            .Select(index => new DocxTableRow(
                [new DocxTableCell("Nested" + index.ToString(CultureInfo.InvariantCulture), [CreateDocxLayoutParagraph("Nested" + index.ToString(CultureInfo.InvariantCulture), 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty)],
                10d))
            .ToArray();
        DocxTable nestedTable = new(null, [60d], nestedRows);
        var earlyCell = new DocxTableCell(string.Empty, [earlyBefore, earlyAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(earlyBefore),
                new DocxPageBreakElement("earlyBreak", "page"),
                new DocxParagraphElement(earlyAfter)
            ]
        };
        var nestedCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        DocxTable table = new(null, [60d, 90d], [new DocxTableRow([earlyCell, nestedCell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table, nestedTable]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        TestAssert.Equal(2, rowFragments.Length);
        TestAssert.Equal(0, rowFragments[0].FragmentIndex);
        TestAssert.Equal(1, rowFragments[1].FragmentIndex);
        TestAssert.Equal("CellPageBreak", rowFragments[0].FragmentReason);
        TestAssert.Equal("CellPageBreak", rowFragments[1].FragmentReason);

        DocxTableRowLayout[] firstFragmentNestedRows = rowFragments[0].Cells[1].NestedRows.ToArray();
        DocxTableRowLayout[] secondFragmentNestedRows = rowFragments[1].Cells[1].NestedRows.ToArray();
        TestAssert.Equal(2, firstFragmentNestedRows.Length);
        TestAssert.Equal(4, secondFragmentNestedRows.Length);
        TestAssert.Equal("Nested1", firstFragmentNestedRows[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Nested2", firstFragmentNestedRows[1].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Nested3", secondFragmentNestedRows[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Nested6", secondFragmentNestedRows[3].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal(0, firstFragmentNestedRows[0].Table.TableIndex);
        TestAssert.Equal(0, secondFragmentNestedRows[0].Table.TableIndex);
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
        TestAssert.Equal(cellLayout.Y + cellLayout.Height - 14d, cellLayout.TextLines[0].BaselineY);
    }

    public static void DocxTableLayoutStageDoesNotInventHorizontalCellPadding()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("Flush", 12d, null, false, false, false, null, null)],
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
        var cell = new DocxTableCell("Flush", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 30d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Flush".EnumerateRunes().Select(rune => rune.Value));

        DocxTableCellLayout cellLayout = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(cellLayout.X, cellLayout.TextLines[0].X);
        TestAssert.Equal(cellLayout.Y + cellLayout.Height - 12d, cellLayout.TextLines[0].BaselineY);
    }

    public static void DocxTableLayoutStageStartsTextInsideVisibleCellBorder()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("Bordered", 12d, null, false, false, false, null, null)],
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
        var borders = new[]
        {
            new DocxTableCellBorder("left", "single", "auto", "4"),
            new DocxTableCellBorder("right", "single", "auto", "4")
        };
        var cell = new DocxTableCell("Bordered", [paragraph], null, null, null, null, borders, DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 30d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Bordered".EnumerateRunes().Select(rune => rune.Value));

        DocxTableCellLayout cellLayout = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(cellLayout.X + 0.24d, cellLayout.TextLines[0].X);
    }

    public static void DocxTableLayoutStageDoesNotInventDefaultRowMinimumForAutoRows()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
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
            null);
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Height);
    }

    public static void DocxTableLayoutStageQuantizesAutoLineHeightToTwips()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("Alpha Beta", 9d, null, false, false, false, null, null)],
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
        var cell = new DocxTableCell("Alpha Beta", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [30d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTextLineLayout[] lines = new DocxLayoutEngine()
            .Create(document, new FractionalLineHeightTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single()
            .TextLines
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(12d, Math.Round(lines[0].BaselineY - lines[1].BaselineY, 6));
    }

    public static void DocxTableLayoutStageIncludesCollapsedHorizontalBorderAdvanceForContentRows()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
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
            null);
        var borders = new[]
        {
            new DocxTableCellBorder("bottom", "single", "auto", "4")
        };
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, borders, DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10.48d, row.Height);
    }

    public static void DocxTableLayoutStageLetsTablePropertyExceptionRowsUseContentHeight()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
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
            null);
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var rowExceptionMargins = new DocxTableCellMargins(0d, null, 0d, null, "0", null, "0", null);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null, TablePropertyExceptionCellMargins: rowExceptionMargins)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Height);
    }

    public static void DocxTableLayoutStageExpandsAtLeastRowsByAuthoredTopCellMargin()
    {
        var defaultCell = new DocxTableCell("A", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var topMarginCell = new DocxTableCell(
            "B",
            [],
            null,
            null,
            null,
            null,
            [],
            new DocxTableCellMargins(12d, null, null, null, "240", null, null, null));
        var table = new DocxTable(null, [40d, 40d], [new DocxTableRow([defaultCell, topMarginCell], 30d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(42d, row.Height);
    }

    public static void DocxTableLayoutStageIncludesParagraphBeforeSpacingInCellHeight()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            12d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(22d, row.Height);
    }

    public static void DocxTableLayoutStageConsumesEmptyCellParagraphLineBox()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("A", 10d, 10d);
        var empty = new DocxParagraph(
            [],
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
            null);
        DocxParagraph second = CreateDocxLayoutParagraph("B", 10d, 10d);
        var cell = new DocxTableCell(string.Empty, [first, empty, second], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(30d, row.Height);
        TestAssert.Equal(2, row.Cells[0].TextLines.Count);
        TestAssert.Equal(20d, Math.Round(row.Cells[0].TextLines[0].BaselineY - row.Cells[0].TextLines[1].BaselineY, 3));
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

    public static void DocxTableLayoutStageSplitsOverwideCellTokensAtSafeCharacterBoundaries()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABCDEFGHIJ", 10d, null, false, false, false, null, null)],
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
        var cell = new DocxTableCell("ABCDEFGHIJ", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(
            null,
            [16d],
            [new DocxTableRow([cell], 10d)],
            PreferredWidthPoints: 16d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.True(cellLayout.TextLines.Count >= 3, "Expected an overwide table-cell token to split into fitting line fragments.");
        TestAssert.Equal("ABC", cellLayout.TextLines[0].Text);
        TestAssert.Equal("DEF", cellLayout.TextLines[1].Text);
        TestAssert.Equal("GHI", cellLayout.TextLines[2].Text);
        TestAssert.True(cellLayout.TextLines[0].EndsWithIntraTokenBreak, "The layout model should preserve that this line ended inside an overwide token.");
        TestAssert.True(cellLayout.TextLines[1].EndsWithIntraTokenBreak, "Subsequent split fragments should keep their intra-token break reason.");
        TestAssert.True(cellLayout.TextLines[2].EndsWithIntraTokenBreak, "Every non-final split fragment should keep its intra-token break reason.");
        TestAssert.True(!cellLayout.TextLines.Last().EndsWithIntraTokenBreak, "The final token fragment should not be marked as a line-end intra-token break.");
        TestAssert.True(cellLayout.TextLines.All(line => line.Width <= cellLayout.Width + 0.001d), "Split token fragments should stay inside the cell frame.");
    }

    public static void DocxTextEmissionOmitsTerminalSpacesAfterIntraTokenBreaks()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        string familyName = font.Value.Resolution.FamilyName;
        var run = new DocxTextRun("ABCDEFGHIJ", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [run],
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
        var cell = new DocxTableCell("ABCDEFGHIJ", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(
            null,
            [16d],
            [new DocxTableRow([cell], 10d)],
            PreferredWidthPoints: 16d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        var renderer = new DocxRenderer(new SingleResolutionFontResolver(font.Value.Resolution));

        DocxTextEmissionSnapshot snapshot = renderer.InspectTextEmission(document);
        DocxTextEmissionLineSnapshot[] splitLines = snapshot.Lines
            .Where(line => !line.IsStaticStory && line.EndsWithIntraTokenBreak)
            .ToArray();

        TestAssert.True(splitLines.Length >= 1, "Expected at least one private-safe emission line marked as an intra-token split.");
        TestAssert.True(splitLines.All(line => line.TerminalSpaceSegmentCount == 0), "Intra-token split prefixes should not synthesize standalone terminal-space PDF operations.");
    }

    public static void DocxTextEmissionSplitsDashPunctuationIntoOfficeLikeOperations()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        string familyName = font.Value.Resolution.FamilyName;
        var run = new DocxTextRun("word-break", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [run],
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
        var renderer = new DocxRenderer(new SingleResolutionFontResolver(font.Value.Resolution));

        DocxTextEmissionLineSnapshot line = renderer.InspectTextEmission(document).Lines.Single(line => !line.IsStaticStory);
        int[] visibleLengths = line.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .Select(segment => segment.TextLength)
            .ToArray();
        double[] visibleStarts = line.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .Select(segment => segment.X)
            .ToArray();

        TestAssert.Equal(3, visibleLengths.Length);
        TestAssert.Equal(4, visibleLengths[0]);
        TestAssert.Equal(1, visibleLengths[1]);
        TestAssert.Equal(5, visibleLengths[2]);
        TestAssert.True(visibleStarts[0] < visibleStarts[1] && visibleStarts[1] < visibleStarts[2], "Dash-punctuation text operations should keep increasing layout origins.");
    }

    public static void DocxTableLayoutDoesNotKeepWholeTableTogetherByDefault()
    {
        var intro = new DocxParagraph(
            [new DocxTextRun("Intro", 10d, null, false, false, false, null, null)],
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
            null);
        var first = new DocxParagraph(
            [new DocxTextRun("First", 10d, null, false, false, false, null, null)],
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
            null);
        var second = new DocxParagraph(
            [new DocxTextRun("Second", 10d, null, false, false, false, null, null)],
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
            null);
        var firstCell = new DocxTableCell("First", [first], null, null, null, null, [], DocxTableCellMargins.Empty);
        var secondCell = new DocxTableCell("Second", [second], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [50d], [
            new DocxTableRow([firstCell], 20d),
            new DocxTableRow([secondCell], 20d)
        ]);
        var document = new DocxDocument(
            100d,
            65d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(intro), new DocxTableElement(table)],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageHonorsExactRowHeightRule()
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
        var table = new DocxTable(null, [34d], [new DocxTableRow([cell], 10d, HeightValue: "200", HeightRuleValue: "exact")]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "First Second".EnumerateRunes().Select(rune => rune.Value));

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, embedded)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Height);
        TestAssert.Equal("200", row.HeightValue ?? string.Empty);
        TestAssert.Equal("exact", row.HeightRuleValue ?? string.Empty);
        TestAssert.Equal(10d, row.DeclaredHeightPoints ?? 0d);
    }

    public static void DocxTableLayoutStageDoesNotExpandExactRowsForCollapsedBorders()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
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
            null);
        var borders = new[]
        {
            new DocxTableCellBorder("bottom", "single", "auto", "4")
        };
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, borders, DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 10d, HeightValue: "200", HeightRuleValue: "exact")]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Height);
    }

    public static void DocxLayoutSnapshotReportsPublicSafeCounts()
    {
        var margins = new DocxTableCellMargins(2d, 3d, 4d, 5d, "40", "60", "80", "100");
        var paragraph = new DocxParagraph(
            [new DocxTextRun("private text is not exposed", 11d, null, false, false, false, null, null)],
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
        var cell = new DocxTableCell(
            "private text is not exposed",
            [paragraph],
            "D9EAF7",
            "clear",
            null,
            "center",
            [new DocxTableCellBorder("top", "single", "000000", "8")],
            margins,
            PreferredWidthPoints: 42d,
            PreferredWidthValue: "840",
            PreferredWidthType: "dxa",
            GridSpan: 2,
            GridSpanValue: "2",
            ConditionalFormat: new DocxTableCellConditionalFormat("100000000000", true, "1", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null),
            HasVerticalMerge: true,
            VerticalMergeValue: "restart");
        DocxTable table = new(
            null,
            [40d, 40d],
            [new DocxTableRow([cell], 20d, IsHeader: true, HeaderValue: "1", HeightValue: "400", HeightRuleValue: "atLeast")],
            PreferredWidthPoints: 84d,
            PreferredWidthValue: "1680",
            PreferredWidthType: "dxa",
            IndentPoints: 6d,
            CellSpacingPoints: 1d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayoutSnapshot snapshot = new DocxRenderer().InspectLayout(document);

        TestAssert.Equal(1, snapshot.Tables.Count);
        DocxTableSnapshot tableSnapshot = snapshot.Tables.Single();
        TestAssert.Equal(0, tableSnapshot.TableIndex);
        TestAssert.Equal(0, tableSnapshot.SourceBlockIndex);
        TestAssert.Equal(0, tableSnapshot.PageStartIndex);
        TestAssert.Equal(0, tableSnapshot.PageEndIndex);
        TestAssert.Equal(1, tableSnapshot.RowCount);
        TestAssert.Equal(1, tableSnapshot.LaidOutRowCount);
        TestAssert.Equal(1, tableSnapshot.HeaderRowLayoutCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredHeaderRowCount);
        TestAssert.Equal(2, tableSnapshot.GridColumnCount);
        TestAssert.Equal(80d, tableSnapshot.GridColumnsWidthSum);
        TestAssert.True(tableSnapshot.HasExplicitGrid, "Snapshot should distinguish authored table grids from inferred grids.");
        TestAssert.Equal(2, tableSnapshot.ResolvedColumnWidths.Count);
        TestAssert.Equal(42d, tableSnapshot.ResolvedColumnWidths[0]);
        TestAssert.Equal(42d, tableSnapshot.ResolvedColumnWidths[1]);
        TestAssert.Equal(84d, tableSnapshot.ResolvedTableWidth);
        TestAssert.Equal(84d, tableSnapshot.PreferredWidthPoints ?? 0d);
        TestAssert.Equal("1680", tableSnapshot.PreferredWidthValue ?? string.Empty);
        TestAssert.Equal("dxa", tableSnapshot.PreferredWidthType ?? string.Empty);
        TestAssert.Equal(6d, tableSnapshot.IndentPoints ?? 0d);
        TestAssert.Equal(1d, tableSnapshot.CellSpacingPoints ?? 0d);
        TestAssert.Equal(1, tableSnapshot.DeclaredHeightRowCount);
        TestAssert.Equal(0, tableSnapshot.ExactHeightRowCount);
        TestAssert.Equal(1, tableSnapshot.AtLeastHeightRowCount);
        TestAssert.Equal(0, tableSnapshot.CantSplitRowCount);
        TestAssert.Equal(0, tableSnapshot.FragmentedRowCount);
        TestAssert.Equal(0, tableSnapshot.FragmentedRowLayoutCount);
        TestAssert.Equal(1, tableSnapshot.MaxRowFragmentCount);
        TestAssert.True(tableSnapshot.HasVerticalMerge, "Snapshot should expose vertical-merge presence without cell text.");
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeCellCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeRestartCellCount);
        TestAssert.Equal(0, tableSnapshot.AuthoredVerticalMergeContinuationCellCount);
        TestAssert.Equal(0, tableSnapshot.LaidOutVerticalMergeContinuationCellCount);
        TestAssert.Equal(0, tableSnapshot.MissingVerticalMergeOwnerCellCount);
        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(1, snapshot.Pages[0].ItemCount);
        TestAssert.Equal(1, snapshot.Pages[0].TableRowCount);
        TestAssert.Equal(0, snapshot.Pages[0].TextLineCount);
        TestAssert.True(snapshot.Pages[0].VerticalUsed >= 20d, "Snapshot should report vertical consumption from laid-out table rows.");
        TestAssert.Equal(Math.Round(snapshot.Pages[0].VerticalUsed, 6), Math.Round(snapshot.Pages[0].TableRowHeightSum, 6));
        TestAssert.Equal(0d, snapshot.Pages[0].TextLineHeightSum);
        TestAssert.Equal(0d, snapshot.Pages[0].InlineImageHeightSum);
        DocxLayoutItemSnapshot row = snapshot.Pages[0].Items.Single();
        TestAssert.Equal("TableRow", row.Kind);
        TestAssert.Equal(1, row.CellCount);
        TestAssert.True(row.TextLength > 0, "Snapshot should expose text length only, not the text itself.");
        TestAssert.Equal(1, snapshot.Pages[0].TableRows.Count);
        DocxTableRowSnapshot tableRow = snapshot.Pages[0].TableRows[0];
        TestAssert.Equal(0, tableRow.TableIndex);
        TestAssert.Equal(0, tableRow.SourceBlockIndex);
        TestAssert.Equal(0, tableRow.PageRowIndex);
        TestAssert.Equal(0, tableRow.RowIndex);
        TestAssert.Equal(1, tableRow.TableRowCount);
        TestAssert.Equal("None", tableRow.FragmentReason);
        TestAssert.Equal(2, tableRow.GridColumnCount);
        TestAssert.Equal(80d, tableRow.GridColumnsWidthSum);
        TestAssert.True(tableRow.HasExplicitGrid, "Row snapshot should carry the table grid provenance.");
        TestAssert.Equal(2, tableRow.ResolvedColumnWidths.Count);
        TestAssert.Equal(42d, tableRow.ResolvedColumnWidths[0]);
        TestAssert.Equal(42d, tableRow.ResolvedColumnWidths[1]);
        TestAssert.Equal(84d, tableRow.ResolvedTableWidth);
        TestAssert.Equal(84d, tableRow.PreferredTableWidthPoints ?? 0d);
        TestAssert.Equal(20d, tableRow.DeclaredHeightPoints ?? 0d);
        TestAssert.Equal("400", tableRow.HeightValue ?? string.Empty);
        TestAssert.Equal("atLeast", tableRow.HeightRuleValue ?? string.Empty);
        TestAssert.True(tableRow.IsHeader, "Snapshot should expose header-row status without text content.");
        TestAssert.Equal("1", tableRow.HeaderValue ?? string.Empty);
        TestAssert.True(tableRow.HasTablePropertyExceptionCellMargins == false, "Snapshot should expose row property-exception presence without document text.");
        TestAssert.Equal(1, tableRow.CellCount);
        TestAssert.True(tableRow.TextLength > 0, "Snapshot should report table row text length only.");
        DocxTableCellSnapshot tableCell = tableRow.Cells.Single();
        TestAssert.Equal(0, tableCell.CellIndex);
        TestAssert.Equal(2, tableCell.GridSpan);
        TestAssert.Equal("2", tableCell.GridSpanValue ?? string.Empty);
        TestAssert.Equal(42d, tableCell.PreferredWidthPoints ?? 0d);
        TestAssert.Equal("dxa", tableCell.PreferredWidthType ?? string.Empty);
        TestAssert.Equal("center", tableCell.VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal(2d, tableCell.MarginTopPoints ?? 0d);
        TestAssert.Equal(3d, tableCell.MarginRightPoints ?? 0d);
        TestAssert.Equal(4d, tableCell.MarginBottomPoints ?? 0d);
        TestAssert.Equal(5d, tableCell.MarginLeftPoints ?? 0d);
        TestAssert.Equal(5d, tableCell.ResolvedPaddingLeftPoints);
        TestAssert.Equal(2d, tableCell.ResolvedPaddingTopPoints);
        TestAssert.Equal(3d, tableCell.ResolvedPaddingRightPoints);
        TestAssert.Equal(4d, tableCell.ResolvedPaddingBottomPoints);
        TestAssert.Equal(tableCell.X + tableCell.ResolvedPaddingLeftPoints, tableCell.ContentBoxX);
        TestAssert.Equal(tableCell.Y + tableCell.ResolvedPaddingBottomPoints, tableCell.ContentBoxY);
        TestAssert.Equal(tableCell.Width - tableCell.ResolvedPaddingLeftPoints - tableCell.ResolvedPaddingRightPoints, tableCell.ContentBoxWidth);
        TestAssert.Equal(tableCell.Height - tableCell.ResolvedPaddingTopPoints - tableCell.ResolvedPaddingBottomPoints, tableCell.ContentBoxHeight);
        TestAssert.True(tableCell.FirstTextLineX is not null, "Snapshot should expose private-safe cell text x-position.");
        TestAssert.True((tableCell.FirstTextLineX ?? 0d) >= tableCell.ContentBoxX, "Cell text should be positioned inside the resolved content box.");
        TestAssert.True(tableCell.FirstBaselineY is not null, "Snapshot should expose private-safe first baseline.");
        TestAssert.Equal(11d, tableCell.FirstBaselineInset);
        TestAssert.True(tableCell.LastBaselineY is not null, "Snapshot should expose private-safe last baseline.");
        TestAssert.Equal(1, tableCell.BorderCount);
        TestAssert.True(tableCell.HasFill, "Snapshot should expose fill presence without the fill value.");
        TestAssert.True(tableCell.HasShadingValue, "Snapshot should expose shading presence without the shading color.");
        TestAssert.True(tableCell.HasConditionalFormat, "Snapshot should expose conditional-format presence without document text.");
        TestAssert.True(tableCell.HasVerticalMerge, "Snapshot should expose vertical-merge presence without document text.");
        TestAssert.Equal("restart", tableCell.VerticalMergeValue ?? string.Empty);
        TestAssert.Equal("OwnCell", tableCell.VisualOwnership);
        TestAssert.True(tableCell.VerticalMergeOwnerRowIndex is null, "Restart cells should not report an owner row.");
        TestAssert.True(tableCell.VerticalMergeOwnerGridColumnIndex is null, "Restart cells should not report an owner grid column.");
    }

    public static void DocxLayoutSnapshotNormalizesPlainTableCellText()
    {
        DocxTable table = CreateSingleCellTable("Plain 123", 20d);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayoutSnapshot snapshot = new DocxRenderer().InspectLayout(document);

        DocxTableRowSnapshot row = snapshot.Pages[0].TableRows.Single();
        DocxTableCellSnapshot cell = row.Cells.Single();
        TestAssert.Equal(9, row.TextLength);
        TestAssert.Equal(9, cell.TextLength);
        TestAssert.Equal(1, cell.ParagraphCount);
        TestAssert.Equal(1, cell.SpaceCharacterCount);
        TestAssert.Equal(3, cell.DigitCharacterCount);
        TestAssert.Equal(5, cell.LongestBreakableTokenLength);
        TestAssert.Equal(1, cell.BodyElementCount);
        TestAssert.Equal(0, cell.ManualBreakElementCount);
        TestAssert.Equal(0, cell.PageBreakElementCount);
        TestAssert.Equal(0, cell.NestedTableElementCount);
    }

    public static void DocxLayoutSnapshotReportsTableCellBodyFlowCounts()
    {
        DocxParagraph paragraph = CreateDocxLayoutParagraph("Flow", 10d, 10d);
        DocxTable nestedTable = CreateSingleCellTable("Nested", 12d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(paragraph),
                new DocxManualBreakElement("runBreak", "column"),
                new DocxPageBreakElement("runBreak", "page"),
                new DocxTableElement(nestedTable)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table, nestedTable]);

        DocxLayoutSnapshot snapshot = new DocxRenderer().InspectLayout(document);

        DocxTableCellSnapshot cellSnapshot = snapshot.Pages[0].TableRows.Single().Cells.Single();
        TestAssert.Equal(4, cellSnapshot.BodyElementCount);
        TestAssert.Equal(1, cellSnapshot.ManualBreakElementCount);
        TestAssert.Equal(1, cellSnapshot.PageBreakElementCount);
        TestAssert.Equal(1, cellSnapshot.NestedTableElementCount);
    }

    public static void DocxLayoutSnapshotReportsOfficeTableCellBaselineFixture()
    {
        string input = Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "docx-ladder-03-table-cell-baseline.docx");
        input = Path.GetFullPath(input);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxLayoutSnapshot snapshot = new DocxRenderer().InspectLayout(document);

        DocxTableRowSnapshot[] rows = snapshot.Pages.SelectMany(page => page.TableRows).ToArray();
        TestAssert.Equal(5, rows.Length);
        TestAssert.True(rows.All(row => row.CellCount == 4), "The Office-authored baseline fixture should keep a stable 5x4 table.");
        TestAssert.Equal(8d, rows[0].Cells[0].FirstBaselineInset);
        TestAssert.Equal(11d, rows[1].Cells[0].FirstBaselineInset);
        TestAssert.Equal(16d, rows[2].Cells[0].FirstBaselineInset);
        TestAssert.Equal(11d, rows[3].Cells[0].FirstBaselineInset);
        TestAssert.Equal(16d, rows[4].Cells[0].FirstBaselineInset);
        TestAssert.True(rows[0].Cells[0].MarginTopPoints is null || rows[0].Cells[0].MarginTopPoints == 0d, "Zero top padding may serialize as absent or zero.");
        TestAssert.Equal(6d, rows[3].Cells[0].MarginTopPoints ?? 0d);
        TestAssert.Equal(6d, rows[4].Cells[0].MarginTopPoints ?? 0d);
    }

    public static void DocxLayoutSnapshotReportsTableCellParagraphIndexes()
    {
        DocxParagraph first = CreateDocxLayoutParagraph("First line", 10d, 12d);
        DocxParagraph second = CreateDocxLayoutParagraph("Second line", 10d, 12d);
        var cell = new DocxTableCell(string.Empty, [first, second], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], 50d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);

        DocxTableCellSnapshot cellSnapshot = snapshot.Pages[0].TableRows.Single().Cells.Single();
        TestAssert.Equal(2, cellSnapshot.ParagraphCount);
        DocxTextLineLayout[] lines = layout.Pages[0].Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells.Single()
            .TextLines.ToArray();
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(0, lines[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, lines[0].SourceLineIndex ?? -1);
        TestAssert.Equal(1, lines[1].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, lines[1].SourceLineIndex ?? -1);
    }

    public static void DocxTextEmissionSnapshotReportsTableCellParagraphIndexes()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        DocxParagraph first = CreateDocxLayoutParagraph("First", 10d, 12d);
        DocxParagraph second = CreateDocxLayoutParagraph("Second", 10d, 12d);
        var cell = new DocxTableCell(string.Empty, [first, second], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], 50d)]);
        DocxDocument document = CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        var renderer = new DocxRenderer(new SingleResolutionFontResolver(font.Value.Resolution));

        DocxTextEmissionLineSnapshot[] lines = renderer.InspectTextEmission(document).Lines
            .Where(line => !line.IsStaticStory)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines.All(line => line.SourceBlockIndex is null), "Table-cell line snapshots should not pretend to be top-level body blocks.");
        TestAssert.Equal(0, lines[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, lines[1].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, lines[0].SourceLineIndex ?? -1);
        TestAssert.Equal(0, lines[1].SourceLineIndex ?? -1);
        TestAssert.True(lines[0].Segments.All(segment => segment.SourceParagraphIndex == 0), "First table-cell paragraph ownership should survive to emission segments.");
        TestAssert.True(lines[1].Segments.All(segment => segment.SourceParagraphIndex == 1), "Second table-cell paragraph ownership should survive to emission segments.");
    }

    public static void DocxLayoutSnapshotReportsPrivateSafeSourceLineIndexes()
    {
        var first = new DocxParagraph(
            [new DocxTextRun("Alpha Beta Gamma", 11d, null, false, false, false, null, null)],
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
        var second = new DocxParagraph(
            [new DocxTextRun("Delta Epsilon", 11d, null, false, false, false, null, null)],
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
            null)
        {
            Indent = new DocxParagraphIndent(120d, null, null, null, null, null, null, null)
        };
        DocxDocument document = CreateLayoutTestDocument(
            [new DocxParagraphElement(first), new DocxParagraphElement(second)],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));
        DocxLayoutItemSnapshot[] textLines = snapshot.Pages[0].Items
            .Where(item => item.Kind == "TextLine")
            .ToArray();

        TestAssert.Equal(3, textLines.Length);
        TestAssert.Equal(0, textLines[0].SourceBlockIndex ?? -1);
        TestAssert.Equal(0, textLines[0].SourceLineIndex ?? -1);
        TestAssert.Equal(1, textLines[1].SourceBlockIndex ?? -1);
        TestAssert.Equal(0, textLines[1].SourceLineIndex ?? -1);
        TestAssert.Equal(1, textLines[2].SourceBlockIndex ?? -1);
        TestAssert.Equal(1, textLines[2].SourceLineIndex ?? -1);
        TestAssert.True(textLines.All(line => line.TextLength > 0), "Snapshot source indexes must not expose line text.");
    }

    public static void DocxLayoutSnapshotReportsLineHeightProfileFacts()
    {
        var label = new DocxListLabel(
            "1",
            "decimal",
            "%1.",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, null, new DocxRunFonts(null, null, null, null, null, null, null, null)));
        var autoSpacing = new DocxParagraphSpacing(null, null, null, null, null, null, "276", "auto", null);
        var flooredList = new DocxParagraph(
            [new DocxTextRun("Floored", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.15d,
            null,
            autoSpacing,
            DocxParagraphKeepRules.Empty,
            label);
        var listWithoutBeforeSpacing = new DocxParagraph(
            [new DocxTextRun("No before", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.15d,
            null,
            autoSpacing,
            DocxParagraphKeepRules.Empty,
            label);
        var defaultAutoList = new DocxParagraph(
            [new DocxTextRun("Default auto", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.2d,
            null,
            new DocxParagraphSpacing(null, null, null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            label);
        var plainParagraph = new DocxParagraph(
            [new DocxTextRun("Plain", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.15d,
            null,
            autoSpacing,
            DocxParagraphKeepRules.Empty,
            null);
        DocxDocument document = CreateLayoutTestDocument(
            [new DocxParagraphElement(flooredList), new DocxParagraphElement(listWithoutBeforeSpacing), new DocxParagraphElement(defaultAutoList), new DocxParagraphElement(plainParagraph)],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));
        DocxLayoutItemSnapshot[] textLines = snapshot.Pages[0].Items
            .Where(item => item.Kind == "TextLine")
            .ToArray();

        TestAssert.Equal(4, textLines.Length);
        TestAssert.Equal(10d, textLines[0].SingleLineHeightPoints ?? 0d);
        TestAssert.Equal(0d, textLines[0].PendingAfterSpacingPoints ?? -1d);
        TestAssert.Equal(6d, textLines[0].ParagraphBeforeSpacingPoints ?? 0d);
        TestAssert.Equal(0d, textLines[0].ParagraphAfterSpacingPoints ?? -1d);
        TestAssert.True(textLines[0].ContextualSpacingSuppressed == false, "First paragraph should report that contextual spacing suppression did not apply.");
        TestAssert.Equal(1.19d, textLines[0].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(Math.Abs((textLines[0].LineHeightPoints ?? 0d) - 11.9d) < 0.0001d, "Effective line height should be the measured single-line height multiplied by the effective factor.");
        TestAssert.True(textLines[0].LineSpacingFactorFloorApplied == true, "Positive before-spacing list paragraphs should report the Word-compatible auto-line floor.");
        TestAssert.Equal(0d, textLines[1].PendingAfterSpacingPoints ?? -1d);
        TestAssert.Equal(0d, textLines[1].ParagraphBeforeSpacingPoints ?? -1d);
        TestAssert.Equal(1.15d, textLines[1].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(textLines[1].LineSpacingFactorFloorApplied == false, "Lists without positive before spacing should not report the floor.");
        TestAssert.Equal(1.2d, textLines[2].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(textLines[2].LineSpacingFactorFloorApplied == false, "Missing w:line list paragraphs should use the Word default auto factor without the explicit-line floor.");
        TestAssert.Equal(1.15d, textLines[3].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(textLines[3].LineSpacingFactorFloorApplied == false, "Non-list paragraphs should not report the list floor.");
    }

    public static void DocxLayoutSnapshotReportsListLabelLineHeightMetricCandidates()
    {
        var label = new DocxListLabel(
            "*",
            "bullet",
            "*",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(
                10d,
                null,
                false,
                false,
                false,
                null,
                "Label Metrics",
                new DocxRunFonts("Label Metrics", null, null, null, null, null, null, null)));
        var paragraph = new DocxParagraph(
            [new DocxTextRun("Body", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.15d,
            null,
            new DocxParagraphSpacing(null, null, null, null, null, null, "276", "auto", null),
            DocxParagraphKeepRules.Empty,
            label);
        DocxDocument document = CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxLayoutItemSnapshot line = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()))
            .Pages[0]
            .Items
            .Single(item => item.Kind == "TextLine");

        TestAssert.Equal(10d, line.SingleLineHeightPoints ?? 0d);
        TestAssert.Equal(10d, line.ListLabelSingleLineHeightPoints ?? 0d);
        TestAssert.Equal(12d, line.BodyWindowsLineHeightPoints ?? 0d);
        TestAssert.Equal(14d, line.ListLabelWindowsLineHeightPoints ?? 0d);
        TestAssert.Equal(11.9d, Math.Round(line.LineHeightPoints ?? 0d, 2));
        TestAssert.True((line.ListLabelWindowsLineHeightPoints ?? 0d) > (line.BodyWindowsLineHeightPoints ?? 0d), "Snapshot should expose when list-label Windows extents exceed body extents.");
    }

    public static void DocxTextEmissionSnapshotReportsPrivateSafePdfTextState()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        string familyName = font.Value.Resolution.FamilyName;
        var spacedRun = new DocxTextRun("Body", 10d, null, false, false, false, null, familyName, CharacterSpacingPoints: 1.25d)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var spacedParagraph = new DocxParagraph(
            [spacedRun],
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
        var label = new DocxListLabel(
            "1",
            "decimal",
            "%1.",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, familyName, new DocxRunFonts(familyName, null, null, null, null, null, null, null)));
        var numberedRun = new DocxTextRun("Item", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var numberedParagraph = new DocxParagraph(
            [numberedRun],
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
            [new DocxParagraphElement(spacedParagraph), new DocxParagraphElement(numberedParagraph)],
            [spacedParagraph, numberedParagraph],
            []);
        var renderer = new DocxRenderer(new SingleResolutionFontResolver(font.Value.Resolution));

        DocxTextEmissionSnapshot snapshot = renderer.InspectTextEmission(document);

        TestAssert.True(snapshot.LineCount >= 2, $"Text-emission snapshot should expose rendered body lines; got {snapshot.LineCount}.");
        TestAssert.True(snapshot.SegmentCount >= snapshot.LineCount, "Text-emission snapshot should expose line segments.");
        TestAssert.True(snapshot.TerminalSpaceSegmentCount >= 1, "Text-emission snapshot should expose Office-like terminal line-space emission.");
        TestAssert.True(snapshot.NonzeroPdfCharacterSpacingSegmentCount >= 1, "Text-emission snapshot should expose PDF text-state character spacing.");
        TestAssert.True(snapshot.CompensatedCharacterSpacingSegmentCount >= 1, "Text-emission snapshot should expose glyph-positioning compensation.");

        DocxTextEmissionLineSnapshot spacedLine = snapshot.Lines.Single(line => line.SourceBlockIndex == 0 && line.SourceLineIndex == 0);
        DocxTextEmissionSegmentSnapshot spacedSegment = spacedLine.Segments.First(segment => !segment.IsTerminalLineSpace);
        TestAssert.Equal(4, spacedSegment.TextLength);
        TestAssert.Equal(4, spacedSegment.CharacterProfile.LetterCount);
        TestAssert.Equal(0, spacedSegment.CharacterProfile.DigitCount);
        TestAssert.Equal(4, spacedSegment.AdvanceProfile.GlyphCount);
        TestAssert.Equal(3, spacedSegment.AdvanceProfile.GlyphGapCount);
        TestAssert.Equal(4, spacedSegment.GlyphAdvanceSignature.GlyphCount);
        TestAssert.Equal(3, spacedSegment.GlyphAdvanceSignature.GlyphPairCount);
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.UnitsPerEm > 0, "Snapshot should expose the font design grid used by advance signatures.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.AdvanceUnits > 0, "Snapshot should expose private-safe glyph advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceUnits > 0, "Snapshot should expose private-safe glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceUnits > 0, "Snapshot should expose private-safe left-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceUnits > 0, "Snapshot should expose private-safe right-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceMinUnits > 0, "Snapshot should expose the minimum glyph-pair advance.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceMaxUnits >= spacedSegment.GlyphAdvanceSignature.PairAdvanceMinUnits, "Snapshot should expose a stable glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMinUnits > 0, "Snapshot should expose a stable left-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxUnits >= spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMinUnits, "Snapshot should expose a stable left-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMinUnits > 0, "Snapshot should expose a stable right-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMaxUnits >= spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMinUnits, "Snapshot should expose a stable right-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceEm > 0d, "Snapshot should expose normalized glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceEm > 0d, "Snapshot should expose normalized left-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceEm > 0d, "Snapshot should expose normalized right-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceMaxEm >= spacedSegment.GlyphAdvanceSignature.PairAdvanceMinEm, "Snapshot should expose normalized glyph-pair advance ranges.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxEm >= spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMinEm, "Snapshot should expose normalized left-side glyph-pair advance ranges.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMaxEm >= spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMinEm, "Snapshot should expose normalized right-side glyph-pair advance ranges.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.Hash.Length == 16, "Snapshot should expose a fixed-width glyph advance signature hash.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairHash.Length == 16, "Snapshot should expose a fixed-width glyph-pair advance signature hash.");
        TestAssert.True(spacedSegment.AdvanceProfile.NaturalPdfWidth > 0d, "Snapshot should expose natural PDF font advance.");
        TestAssert.True(spacedSegment.AdvanceProfile.UnkernedPdfWidth > 0d, "Snapshot should expose unkerned PDF font advance.");
        TestAssert.True(spacedSegment.AdvanceProfile.RoundedPdfWidth > 0d, "Snapshot should expose rounded PDF width-array advance.");
        TestAssert.True(spacedSegment.AdvanceProfile.PositioningCharacterSpacingGapTotal > 0d, "Snapshot should expose positioned-spacing contribution to the emitted PDF advance.");
        TestAssert.True(Math.Abs(spacedSegment.AdvanceProfile.TextStateCharacterSpacingGapTotal) < 0.0001d, "Authored DOCX run spacing should not become PDF Tc in this guard.");
        TestAssert.True(spacedSegment.AdvanceProfile.PlannedEmittedAdvance > spacedSegment.AdvanceProfile.RoundedPdfWidth, "Snapshot should expose candidate emitted advance after positioning adjustments.");
        TestAssert.True(Math.Abs(spacedSegment.AdvanceProfile.PlannedEmittedAdvance - (
            spacedSegment.AdvanceProfile.RoundedPdfWidth +
            spacedSegment.AdvanceProfile.KerningAdjustmentTotal +
            spacedSegment.AdvanceProfile.PositioningCharacterSpacingGapTotal +
            spacedSegment.AdvanceProfile.TextStateCharacterSpacingGapTotal)) < 0.0001d, "Planned emitted advance should be decomposed into width-array, kerning, positioning, and Tc terms.");
        TestAssert.True(spacedSegment.AdvanceProfile.UniformResidualPerGap is not null, "Multi-glyph operations should expose residual per glyph gap.");
        TestAssert.True(spacedSegment.AdvanceProfile.RoundedResidualPerGap is not null, "Multi-glyph operations should expose rounded-PDF residual per glyph gap.");
        TestAssert.True(spacedSegment.AdvanceProfile.PlannedEmittedResidualPerGap is not null, "Multi-glyph operations should expose planned emitted residual per glyph gap.");
        TestAssert.True(Math.Abs(spacedSegment.LayoutCharacterSpacing - 1.25d) < 0.0001d, "Snapshot should preserve authored run character spacing.");
        TestAssert.True(Math.Abs(spacedSegment.PdfCharacterSpacing) < 0.0001d, "Normal DOCX run spacing should stay in positioned glyph advances.");
        TestAssert.Equal("None", spacedSegment.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(spacedSegment.PositioningCharacterSpacing - 1.25d) < 0.0001d, "Snapshot should expose the resulting glyph-positioning spacing.");
        TestAssert.True(spacedSegment.CompensatePdfCharacterSpacing, "Run spacing should be marked as compensated when positioned glyph advances carry the spacing.");
        TestAssert.Equal(1, spacedLine.TerminalSpaceSegmentCount);

        DocxTextEmissionLineSnapshot numberedLine = snapshot.Lines.First(line => line.SourceBlockIndex == 1);
        DocxTextEmissionSegmentSnapshot labelSegment = numberedLine.Segments.First(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d);
        TestAssert.Equal("ListLabel", labelSegment.Role);
        TestAssert.Equal(1, labelSegment.TextLength);
        TestAssert.Equal(1, labelSegment.CharacterProfile.DigitCount);
        TestAssert.Equal(0, labelSegment.CharacterProfile.LetterCount);
        TestAssert.Equal(1, labelSegment.AdvanceProfile.GlyphCount);
        TestAssert.Equal(0, labelSegment.AdvanceProfile.GlyphGapCount);
        TestAssert.Equal(1, labelSegment.GlyphAdvanceSignature.GlyphCount);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.GlyphPairCount);
        TestAssert.True(labelSegment.GlyphAdvanceSignature.AdvanceUnits > 0, "Single-glyph labels should still expose an advance signature.");
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairAdvanceUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairRightAdvanceUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairAdvanceMinUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairAdvanceMaxUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMinUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMinUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMaxUnits);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairAdvanceEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairRightAdvanceEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairAdvanceMinEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairAdvanceMaxEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMinEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMinEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMaxEm);
        TestAssert.True(labelSegment.AdvanceProfile.UniformResidualPerGap is null, "Single-glyph operations should not report a per-gap residual.");
        TestAssert.True(labelSegment.AdvanceProfile.RoundedResidualPerGap is null, "Single-glyph operations should not report a rounded-PDF per-gap residual.");
        TestAssert.True(Math.Abs(labelSegment.PdfCharacterSpacing - 0.04d) < 0.0001d, "Numbering labels should expose their PDF text-state character spacing.");
        TestAssert.Equal("ListLabel", labelSegment.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(labelSegment.PositioningCharacterSpacing) < 0.0001d, "Numbering PDF text-state spacing should not be double-counted in glyph positioning.");
        TestAssert.True(!labelSegment.CompensatePdfCharacterSpacing, "Numbering label spacing is intentionally emitted through PDF text state.");
        TestAssert.True(labelSegment.FontResourceName is not null, "Snapshot should identify the resolved PDF font resource without exposing text.");
        TestAssert.True(numberedLine.Segments.Any(segment => segment.Role == "ListSeparator"), "Numbered lines should distinguish the marker separator from body text.");
        TestAssert.True(numberedLine.Segments.Any(segment => segment.Role == "Text"), "Numbered lines should preserve paragraph text as ordinary text segments.");
    }

    public static void DocxTextEmissionDoesNotApplyNumberedTcToBulletListMarkers()
    {
        (FontResolution Resolution, OpenTypeFont Font)? font = FindUsableInstalledFont();
        if (font is null)
        {
            return;
        }

        string familyName = font.Value.Resolution.FamilyName;
        var label = new DocxListLabel(
            "\uF0B7",
            "bullet",
            "\uF0B7",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, familyName, new DocxRunFonts(familyName, null, null, null, null, null, null, null)));
        var run = new DocxTextRun("Item", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [run],
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
        var renderer = new DocxRenderer(new SingleResolutionFontResolver(font.Value.Resolution));

        DocxTextEmissionLineSnapshot line = renderer.InspectTextEmission(document).Lines.Single();

        TestAssert.True(
            line.Segments.All(segment => Math.Abs(segment.PdfCharacterSpacing) < 0.0001d),
            "Bullet-format list markers should keep PDF Tc at zero; decimal numbering labels remain the separate Tc branch.");
    }

    public static void DocxLayoutSnapshotReportsTableSourceBlockIndexes()
    {
        DocxParagraph before = CreateDocxLayoutParagraph("Before", 10d, 12d);
        DocxTable first = CreateSingleCellTable("first", 20d);
        DocxParagraph middle = CreateDocxLayoutParagraph("Middle", 10d, 12d);
        DocxTable second = CreateSingleCellTable("second", 20d);
        DocxDocument document = CreateLayoutTestDocument(
            [
                new DocxParagraphElement(before),
                new DocxTableElement(first),
                new DocxParagraphElement(middle),
                new DocxTableElement(second)
            ],
            [first, second]);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));

        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(4, snapshot.Pages[0].SourceBlockCount);
        TestAssert.Equal(0, snapshot.Pages[0].FirstSourceBlockIndex ?? -1);
        TestAssert.Equal(3, snapshot.Pages[0].LastSourceBlockIndex ?? -1);
        TestAssert.Equal(4, snapshot.SourceBlocks.Count);
        DocxLayoutSourceBlockSnapshot firstBlock = snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 0);
        TestAssert.Equal("Paragraph", firstBlock.Kind);
        TestAssert.Equal(1, firstBlock.TextLineCount);
        TestAssert.Equal(0, firstBlock.TableRowCount);
        TestAssert.True(firstBlock.VerticalTop >= firstBlock.VerticalBottom, "Source block vertical bounds should describe the emitted layout span.");
        TestAssert.True(firstBlock.TextLength > 0, "Source block summary should expose text length only, not text.");
        DocxLayoutSourceBlockSnapshot firstTableBlock = snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 1);
        TestAssert.Equal("Table", firstTableBlock.Kind);
        TestAssert.Equal(0, firstTableBlock.TextLineCount);
        TestAssert.Equal(1, firstTableBlock.TableRowCount);
        TestAssert.Equal(0, firstTableBlock.FirstPageIndex);
        TestAssert.Equal(0, firstTableBlock.LastPageIndex);
        TestAssert.True(firstTableBlock.VerticalTop > firstTableBlock.VerticalBottom, "Table source block bounds should include the row band.");
        TestAssert.Equal("Paragraph", snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 2).Kind);
        TestAssert.Equal("Table", snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 3).Kind);
        TestAssert.Equal(2, snapshot.Tables.Count);
        TestAssert.Equal(0, snapshot.Tables[0].TableIndex);
        TestAssert.Equal(1, snapshot.Tables[0].SourceBlockIndex);
        TestAssert.Equal(1, snapshot.Tables[1].TableIndex);
        TestAssert.Equal(3, snapshot.Tables[1].SourceBlockIndex);
        DocxTableRowSnapshot[] rows = snapshot.Pages.SelectMany(page => page.TableRows).ToArray();
        TestAssert.Equal(2, rows.Length);
        TestAssert.Equal(1, rows.Single(row => row.TableIndex == 0).SourceBlockIndex);
        TestAssert.Equal(3, rows.Single(row => row.TableIndex == 1).SourceBlockIndex);
        DocxLayoutItemSnapshot[] tableItems = snapshot.Pages.SelectMany(page => page.Items)
            .Where(item => item.Kind == "TableRow")
            .ToArray();
        TestAssert.Equal(2, tableItems.Length);
        TestAssert.Equal(1, tableItems[0].SourceBlockIndex ?? -1);
        TestAssert.Equal(3, tableItems[1].SourceBlockIndex ?? -1);
    }

    public static void DocxLayoutSnapshotReportsInlineImageSourceBlockIndexes()
    {
        var image = new DocxInlineImage(24d, 18d, "image/png", [0x89, 0x50, 0x4E, 0x47], "word/media/image1.png");
        DocxParagraph imageParagraph = CreateDocxLayoutParagraph(string.Empty, 10d, 12d) with
        {
            Runs = [],
            Images = [image]
        };
        DocxDocument document = CreateLayoutTestDocument([new DocxParagraphElement(imageParagraph)], []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));

        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(1, snapshot.Pages[0].InlineImageCount);
        TestAssert.Equal(1, snapshot.Pages[0].SourceBlockCount);
        DocxLayoutItemSnapshot imageItem = snapshot.Pages[0].Items.Single(item => item.Kind == "InlineImage");
        TestAssert.Equal(0, imageItem.SourceBlockIndex ?? -1);
        DocxLayoutSourceBlockSnapshot sourceBlock = snapshot.SourceBlocks.Single();
        TestAssert.Equal(0, sourceBlock.SourceBlockIndex);
        TestAssert.Equal("InlineImage", sourceBlock.Kind);
        TestAssert.Equal(1, sourceBlock.InlineImageCount);
        TestAssert.Equal(0, sourceBlock.TextLineCount);
        TestAssert.Equal(0, sourceBlock.TextLength);
        TestAssert.Equal(0, sourceBlock.TableRowCount);
        TestAssert.True(sourceBlock.VerticalTop > sourceBlock.VerticalBottom, "Inline image source block bounds should include the image rectangle.");
    }

    public static void DocxLayoutStageOwnsSelectedStaticHeaderFooterLines()
    {
        DocxParagraph header = new(
            [
                new DocxTextRun("H", 10d, "FF0000", false, false, false, null, "Narrow"),
                new DocxTextRun("{PAGE}", 14d, "0000FF", false, false, false, null, "Narrow")
            ],
            [],
            null,
            DocxTextAlignment.Center,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph footer = new(
            [new DocxTextRun("F", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Right,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = new(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            12d,
            12d,
            "240",
            "240",
            null,
            null,
            null,
            null)
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            },
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [footer]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        DocxTextLineLayout[] staticLines = layout.Pages[0].StaticTextLines.ToArray();
        TestAssert.Equal(2, staticLines.Length);
        TestAssert.Equal("H1", staticLines[0].Text);
        TestAssert.Equal(95d, staticLines[0].X);
        TestAssert.Equal(174d, staticLines[0].BaselineY);
        TestAssert.Equal(2, staticLines[0].Segments.Count);
        TestAssert.Equal(10d, staticLines[0].Segments[0].FontSize ?? 0d);
        TestAssert.Equal(14d, staticLines[0].Segments[1].FontSize ?? 0d);
        TestAssert.Equal("1", staticLines[0].Segments[1].Text);
        TestAssert.Equal("0000FF", staticLines[0].Segments[1].StyleRun.ColorHex ?? string.Empty);
        TestAssert.Equal("F", staticLines[1].Text);
        TestAssert.Equal(185d, staticLines[1].X);
        TestAssert.Equal(14d, staticLines[1].BaselineY);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(2, snapshot.Pages[0].StaticTextLineCount);
        TestAssert.Equal(2, snapshot.Pages[0].StaticItems.Count);
        TestAssert.Equal("StaticHeaderTextLine", snapshot.Pages[0].StaticItems[0].Kind);
        TestAssert.Equal("StaticFooterTextLine", snapshot.Pages[0].StaticItems[1].Kind);
        TestAssert.Equal(2, snapshot.Pages[0].StaticStories.Count);
        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single(story => story.Kind == "Header");
        DocxStaticStoryLayoutSnapshot footerStory = snapshot.Pages[0].StaticStories.Single(story => story.Kind == "Footer");
        TestAssert.True(headerStory.VariantType == "default" && headerStory.TextLineCount == 1 && headerStory.ParagraphCount == 1 && headerStory.TextLength == 2 && headerStory.Items.Single().Kind == "StaticHeaderTextLine", "Static header story snapshots should group private-safe selected header line geometry.");
        TestAssert.True(footerStory.VariantType == "default" && footerStory.TextLineCount == 1 && footerStory.ParagraphCount == 1 && footerStory.TextLength == 1 && footerStory.Items.Single().Kind == "StaticFooterTextLine", "Static footer story snapshots should group private-safe selected footer line geometry.");
        TestAssert.Equal(1, snapshot.Pages[0].TextLineCount);
    }

    public static void DocxLayoutStageSelectsStaticHeaderBodyElements()
    {
        DocxParagraph header = CreateDocxLayoutParagraph("HB", 10d, 10d);
        DocxParagraph body = CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(header)]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));

        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single();
        TestAssert.True(headerStory.Kind == "Header" && headerStory.VariantType == "default" && headerStory.TextLineCount == 1 && headerStory.ParagraphCount == 1 && headerStory.TextLength == 2, "Static header layout should select body elements directly instead of requiring the legacy paragraph map.");
        TestAssert.Equal("StaticHeaderTextLine", snapshot.Pages[0].StaticItems.Single().Kind);
    }

    public static void DocxLayoutStageSummarizesSelectedStaticHeaderFooterVariants()
    {
        DocxParagraph defaultHeader = CreateDocxLayoutParagraph("DH", 10d, 10d);
        DocxParagraph firstHeader = CreateDocxLayoutParagraph("FH", 10d, 10d);
        DocxParagraph evenFooter = CreateDocxLayoutParagraph("EF", 10d, 10d);
        DocxParagraph firstBody = CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph secondBody = CreateDocxLayoutParagraph("Second", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            TitlePage = true,
            EvenAndOddHeaders = true,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [defaultHeader],
                ["first"] = [firstHeader]
            },
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["even"] = [evenFooter]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(firstBody), new DocxPageBreakElement("runBreak", "page"), new DocxParagraphElement(secondBody)],
            [firstBody, secondBody],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));

        TestAssert.Equal(2, snapshot.Pages.Count);
        TestAssert.Equal("first", snapshot.Pages[0].StaticStories.Single().VariantType ?? string.Empty);
        TestAssert.Equal("default", snapshot.Pages[1].StaticStories.Single(story => story.Kind == "Header").VariantType ?? string.Empty);
        TestAssert.Equal("even", snapshot.Pages[1].StaticStories.Single(story => story.Kind == "Footer").VariantType ?? string.Empty);
        TestAssert.True(snapshot.Pages[1].StaticItems.All(item => item.StoryVariantType is "default" or "even"), "Static item snapshots should retain the selected header/footer variant type.");
    }

    public static void DocxLayoutStageOwnsStaticHeaderInlineImages()
    {
        DocxInlineImage headerImage = new(36d, 18d, "image/png", [1, 2, 3], "/word/media/header.png");
        DocxParagraph header = new(
            [],
            [headerImage],
            null,
            DocxTextAlignment.Center,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        DocxInlineImageLayout staticImage = layout.Pages[0].StaticInlineImages.Single();
        TestAssert.True(staticImage.Image == headerImage && staticImage.Width == 36d && staticImage.Height == 18d, "Static header inline images should be first-class page layout items with image geometry.");
        TestAssert.True(staticImage.SourceParagraphIndex == 0 && staticImage.StoryKind == "Header" && staticImage.StoryVariantType == "default", "Static header inline image layout should retain selected-story provenance.");

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(1, snapshot.Pages[0].StaticInlineImageCount);
        DocxLayoutItemSnapshot staticItem = snapshot.Pages[0].StaticItems.Single();
        TestAssert.True(staticItem.Kind == "StaticHeaderInlineImage" && staticItem.StoryVariantType == "default" && staticItem.SourceParagraphIndex == 0, "Static image snapshots should expose private-safe header story ownership.");
        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single();
        TestAssert.True(headerStory.Kind == "Header" && headerStory.TextLineCount == 0 && headerStory.InlineImageCount == 1 && headerStory.ParagraphCount == 1, "Static story snapshots should summarize inline image ownership separately from text line counts.");
    }

    public static void DocxLayoutStageWrapsStaticHeaderLines()
    {
        DocxParagraph header = new(
            [new DocxTextRun("Alpha Beta", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Center,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = new(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            12d,
            12d,
            "240",
            "240",
            null,
            null,
            null,
            null)
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            }
        };
        DocxDocument document = new(
            100d,
            200d,
            10d,
            50d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayout layout = new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer());

        DocxTextLineLayout[] staticLines = layout.Pages[0].StaticTextLines.ToArray();
        TestAssert.Equal(2, staticLines.Length);
        TestAssert.Equal("Alpha ", staticLines[0].Text);
        TestAssert.Equal(15d, staticLines[0].X);
        TestAssert.Equal(178d, staticLines[0].BaselineY);
        TestAssert.Equal(0, staticLines[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, staticLines[0].SourceLineIndex ?? -1);
        TestAssert.True(staticLines[0].IsFirstParagraphLine == true, "The first wrapped static line should carry first-line ownership.");
        TestAssert.Equal("Beta", staticLines[1].Text);
        TestAssert.Equal(20d, staticLines[1].X);
        TestAssert.Equal(166d, staticLines[1].BaselineY);
        TestAssert.Equal(0, staticLines[1].SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, staticLines[1].SourceLineIndex ?? -1);
        TestAssert.True(staticLines[1].IsFirstParagraphLine == false, "Continuation static lines should not look like first paragraph lines.");

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(2, snapshot.Pages[0].StaticTextLineCount);
        TestAssert.Equal(2, snapshot.Pages[0].StaticItems.Count);
        TestAssert.Equal("StaticHeaderTextLine", snapshot.Pages[0].StaticItems[0].Kind);
        TestAssert.Equal(0, snapshot.Pages[0].StaticItems[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, snapshot.Pages[0].StaticItems[0].SourceLineIndex ?? -1);
        TestAssert.True(snapshot.Pages[0].StaticItems[0].IsFirstParagraphLine == true, "The static snapshot should preserve first-line ownership.");
        TestAssert.Equal(1, snapshot.Pages[0].StaticItems[1].SourceLineIndex ?? -1);
        TestAssert.True(snapshot.Pages[0].StaticItems[1].IsFirstParagraphLine == false, "The static snapshot should preserve continuation-line ownership.");
        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single();
        TestAssert.True(headerStory.Kind == "Header" && headerStory.TextLineCount == 2 && headerStory.ParagraphCount == 1 && headerStory.SourceLineCount == 2 && headerStory.FirstParagraphLineCount == 1 && headerStory.Items.Count == 2, "Static story snapshots should summarize wrapped selected header line ownership without exposing text.");
        TestAssert.Equal(0, headerStory.FirstSourceLineIndex ?? -1);
        TestAssert.Equal(1, headerStory.LastSourceLineIndex ?? -1);
    }

    public static void DocxLayoutStageAppliesStaticHeaderParagraphSpacing()
    {
        DocxParagraph first = new(
            [new DocxTextRun("A", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            6d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph second = new(
            [new DocxTextRun("B", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            4d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = new(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            12d,
            12d,
            "240",
            "240",
            null,
            null,
            null,
            null)
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [first, second]
            }
        };
        DocxDocument document = new(
            100d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxTextLineLayout[] staticLines = new DocxLayoutEngine()
            .Create(document, new FamilyWidthTextMeasurer())
            .Pages[0]
            .StaticTextLines
            .ToArray();

        TestAssert.Equal(2, staticLines.Length);
        TestAssert.Equal("A", staticLines[0].Text);
        TestAssert.Equal(178d, staticLines[0].BaselineY);
        TestAssert.Equal("B", staticLines[1].Text);
        TestAssert.Equal(160d, staticLines[1].BaselineY);
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
        TestAssert.Equal(6, CountPdfTextShows(pdf));
    }

    public static void DocxSyntheticHeaderReferenceTypesSelectDefaultWhenNoFirstOrEvenSetting()
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
                  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/header3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
                  <Relationship Id="rIdHeaderEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rIdHeaderDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                  <Relationship Id="rIdHeaderFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header3.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Even header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/header2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Default header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/header3.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="even" r:id="rIdHeaderEven"/>
                      <w:headerReference w:type="default" r:id="rIdHeaderDefault"/>
                      <w:headerReference w:type="first" r:id="rIdHeaderFirst"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream);
            DocxDocument document = new DocxReader().Read(package);
            TestAssert.Equal(3, document.HeaderParagraphsByType.Count);
            TestAssert.Equal(1, document.HeaderParagraphs.Count);
        }

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(4, CountPdfTextShows(pdf));
    }

    public static void DocxReaderPreservesStaticHeaderTablesAsBodyElements()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
                  <Relationship Id="rIdHeaderDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:tbl>
                    <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                    <w:tr>
                      <w:tc><w:p><w:r><w:t>Header table text</w:t></w:r></w:p></w:tc>
                    </w:tr>
                  </w:tbl>
                </w:hdr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeaderDefault"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        IReadOnlyList<DocxBodyElement> headerElements = document.HeaderBodyElementsByType["default"];
        TestAssert.True(headerElements.Count == 1 && headerElements.Single() is DocxTableElement, "Static header parts should preserve tables as body elements instead of dropping them from the paragraph inventory.");
        TestAssert.Equal(0, document.HeaderParagraphsByType["default"].Count);
        TestAssert.Equal(1, document.PageSettings.HeaderBodyElementsByType["default"].Count);

        DocxStructureStorySnapshot headerStory = new DocxRenderer().InspectStructure(document).Stories.Single(story => story.Kind == "Header" && story.VariantType == "default");
        TestAssert.True(headerStory.BlockCount == 1 && headerStory.TableCount == 1 && headerStory.ParagraphCount == 0 && headerStory.TextLength == 17, "Static header structure snapshots should derive counts from body elements, including table-cell paragraphs.");

        DocxFontPlan fontPlan = DocxFontPlan.Create(document, new MapFontResolver([], "Fallback"));
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Header table text"), "Static header table-cell runs should participate in DOCX font planning through block traversal.");
    }

    public static void DocxReaderPreservesHeaderFloatingDrawingsByVariant()
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
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """),
            ["word/_rels/header1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdHeaderImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/header.png"/>
                </Relationships>
                """),
            ["word/header1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:p>
                    <w:r>
                      <w:drawing>
                        <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="1">
                          <wp:extent cx="914400" cy="457200"/>
                          <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                          <wp:positionV relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                          <wp:wrapNone/>
                          <a:graphic>
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                              <pic:pic><pic:blipFill><a:blip r:embed="rIdHeaderImage1"/></pic:blipFill></pic:pic>
                            </a:graphicData>
                          </a:graphic>
                        </wp:anchor>
                      </w:drawing>
                    </w:r>
                  </w:p>
                </w:hdr>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/header.png"] = TestFixtures.CreateRgbPng(1, 1, [32, 64, 96])
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        DocxDocument document = new DocxReader().Read(package);

        DocxFloatingDrawing drawing = document.HeaderFloatingDrawingsByType["default"].Single();
        TestAssert.Equal(1, document.PageSettings.HeaderFloatingDrawingsByType["default"].Count);
        TestAssert.Equal("rIdHeaderImage1", drawing.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/header.png", drawing.Image?.PartName ?? string.Empty);
        TestAssert.Equal(0, drawing.SourceParagraphIndex ?? -1);
        TestAssert.True(drawing.SourceBlockIndex is null, "Header floating drawings should not pretend to belong to a body block.");

        DocxStructureStorySnapshot story = new DocxRenderer().InspectStructure(document).Stories.Single(story => story.Kind == "Header" && story.VariantType == "default");
        TestAssert.True(story.FloatingDrawingCount == 1 && story.ParagraphCount == 1, "Static header story snapshots should expose anchored drawing ownership without rendering it yet.");

        DocxLayoutSnapshot layout = new DocxRenderer().InspectLayout(document);
        DocxFloatingDrawingLayoutSnapshot layoutDrawing = layout.StaticFloatingDrawings.Single();
        TestAssert.True(layout.FloatingDrawings.Count == 0 && layoutDrawing.StoryKind == "Header" && layoutDrawing.StoryVariantType == "default", "Selected header drawings should be laid out in the static drawing stream, not mixed with body floating drawings.");
        TestAssert.Equal(0, layoutDrawing.AnchorPageIndex ?? -1);
        TestAssert.Equal(72d, layoutDrawing.PlacedX ?? 0d);
        TestAssert.Equal(756d, layoutDrawing.PlacedTop ?? 0d);
        TestAssert.Equal("Offset", layoutDrawing.HorizontalPlacementSource ?? string.Empty);
        TestAssert.Equal("Offset", layoutDrawing.VerticalPlacementSource ?? string.Empty);
        TestAssert.Equal(72d, layoutDrawing.ExtentWidthPoints ?? 0d);
        TestAssert.Equal(36d, layoutDrawing.ExtentHeightPoints ?? 0d);

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void DocxSyntheticHeaderFooterDistancesUsePageMarginTokens()
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
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header distance</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer distance</w:t></w:r></w:p></w:ftr>
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
                      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="1440" w:footer="1080"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] leftAlignedBaselines = Regex.Matches(pdf, @"1 0 0 1 36(?:\.\d+)? (?<y>-?\d+(?:\.\d+)?) Tm")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .ToArray();
        TestAssert.True(leftAlignedBaselines.Any(y => y > 704d && y < 720d), "Header baseline should be inset from the raw header-distance top by resolved font ascender metrics.");
        TestAssert.True(leftAlignedBaselines.Any(y => y > 54d && y < 60d), "Footer baseline should be inset from the raw footer-distance bottom by resolved font descender metrics.");
        TestAssert.DoesNotContain("1 0 0 1 36 720 Tm", pdf);
        TestAssert.DoesNotContain("1 0 0 1 36 54 Tm", pdf);
    }

    public static void DocxSyntheticSectionHeadersUsePageLocalGeometry()
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
                  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>Section header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>Section footer</w:t></w:r></w:p></w:ftr>
                """,
            ["word/header2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>Final header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="FF00FF"/></w:rPr><w:t>Final footer</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>First body</w:t></w:r></w:p>
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:headerReference w:type="default" r:id="rIdHeader1"/>
                          <w:footerReference w:type="default" r:id="rIdFooter1"/>
                          <w:pgSz w:w="4000" w:h="4000"/>
                          <w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360" w:header="360" w:footer="360"/>
                          <w:type w:val="nextPage"/>
                        </w:sectPr>
                      </w:pPr>
                    </w:p>
                    <w:p><w:r><w:t>Second body</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader2"/>
                      <w:footerReference w:type="default" r:id="rIdFooter2"/>
                      <w:pgSz w:w="6000" w:h="6000"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="1080" w:footer="1080"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] firstSectionLeftBaselines = ExtractTextBaselinesAtX(pdf, 18d);
        double[] finalSectionLeftBaselines = ExtractTextBaselinesAtX(pdf, 72d);
        TestAssert.True(firstSectionLeftBaselines.Any(y => y > 168d && y < 182d), "First-section header should use the first section header distance and left margin.");
        TestAssert.True(firstSectionLeftBaselines.Any(y => y > 18d && y < 30d), "First-section footer should use the first section footer distance and left margin.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 232d && y < 246d), "Final-section header should use the final section header distance and left margin.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 54d && y < 66d), "Final-section footer should use the final section footer distance and left margin.");
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
        TestAssert.Contains("1 0 1 rg", pdf);
    }

    public static void DocxSyntheticSectionHeadersDoNotBackfillEarlierSectionsFromFinalReferences()
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
                  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
                </Relationships>
                """,
            ["word/header2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>Final header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="FF00FF"/></w:rPr><w:t>Final footer</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>First body</w:t></w:r></w:p>
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:pgSz w:w="4000" w:h="4000"/>
                          <w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360" w:header="360" w:footer="360"/>
                          <w:type w:val="nextPage"/>
                        </w:sectPr>
                      </w:pPr>
                    </w:p>
                    <w:p><w:r><w:t>Second body</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader2"/>
                      <w:footerReference w:type="default" r:id="rIdFooter2"/>
                      <w:pgSz w:w="6000" w:h="6000"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="1080" w:footer="1080"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] firstSectionLeftBaselines = ExtractTextBaselinesAtX(pdf, 18d);
        double[] finalSectionLeftBaselines = ExtractTextBaselinesAtX(pdf, 72d);
        TestAssert.True(!firstSectionLeftBaselines.Any(y => y > 168d && y < 182d), "Final-section header should not backfill the earlier section that omits a header reference.");
        TestAssert.True(!firstSectionLeftBaselines.Any(y => y > 18d && y < 30d), "Final-section footer should not backfill the earlier section that omits a footer reference.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 232d && y < 246d), "Final-section header should still render on its owning section.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 54d && y < 66d), "Final-section footer should still render on its owning section.");
    }

    public static void DocxStaticHeaderRendersMixedRunColorsSeparately()
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
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p>
                    <w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="20"/></w:rPr><w:t>Red</w:t></w:r>
                    <w:r><w:rPr><w:color w:val="0000FF"/><w:sz w:val="40"/></w:rPr><w:t>Blue</w:t></w:r>
                  </w:p>
                </w:hdr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains(" 9.96 Tf", pdf);
        TestAssert.Contains(" 20.04 Tf", pdf);
        TestAssert.Equal(5, CountPdfTextShows(pdf));
    }

    public static void DocxSyntheticFooterPageFieldsUseGeneratedPageNumbers()
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
                  <w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "/><w:r><w:t> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "/></w:p>
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
        TestAssert.True(CountPdfTextShows(pdf) >= 51, "Footer PAGE and NUMPAGES fields should render on each generated page.");
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
                        <w:ind w:startChars="200" w:hangingChars="100"/>
                        <w:sectPr><w:type w:val="continuous"/></w:sectPr>
                      </w:pPr>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins><w:r><w:t>Inserted</w:t></w:r></w:ins>
                      <w:del><w:r><w:t>Deleted</w:t></w:r></w:del>
                      <w:moveFrom><w:r><w:t>Moved</w:t></w:r></w:moveFrom>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE </w:instrText></w:r>
                      <w:r><w:object/></w:r>
                      <w:r><w:drawing><wp:anchor/></w:drawing></w:r>
                      <w:r><w:footnoteReference w:id="2"/></w:r>
                      <w:r><w:endnoteReference w:id="3"/></w:r>
                      <w:r><w:br w:type="column"/></w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:p>
                            <w:pPr><w:keepNext/></w:pPr>
                            <w:r><w:t>Table keep rule</w:t></w:r>
                          </w:p>
                          <w:p><w:r><w:br w:type="column"/></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
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
        TestAssert.Contains("DOCX_UNSUPPORTED_CHARACTER_UNIT_INDENT", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_ENDNOTE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_EQUATION", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_FLOATING_DRAWING", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_FOOTNOTE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_MACRO", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_OLE_OBJECT", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_SECTION_BREAK", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_TRACKED_CHANGES", ids);
        TestAssert.True(diagnostics.All(d => d.Severity == OoxPdfSeverity.Warning && d.PartName == "/word/document.xml"), "Unsupported DOCX diagnostics should be document-scoped warnings.");
    }

    public static void DocxSupportedPageSectionBreaksDoNotEmitUnsupportedSectionDiagnostic()
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
                    <w:p><w:r><w:t>First section</w:t></w:r></w:p>
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:pgSz w:w="4000" w:h="4000"/>
                          <w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/>
                          <w:type w:val="oddPage"/>
                        </w:sectPr>
                      </w:pPr>
                    </w:p>
                    <w:p><w:r><w:t>Second section</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="6000" w:h="6000"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_SECTION_BREAK"), "Supported page-starting paragraph section breaks should not emit a stale unsupported-section diagnostic.");
    }

    public static void DocxSupportedColumnBreakOnlyParagraphDoesNotEmitUnsupportedManualBreakDiagnostic()
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
                    <w:p><w:r><w:t>Left column</w:t></w:r></w:p>
                    <w:p><w:r><w:br w:type="column"/></w:r></w:p>
                    <w:p><w:r><w:t>Right column</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:cols w:num="2"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_MANUAL_BREAK"), "Break-only column paragraphs should be modeled as supported manual-break blocks, not stale unsupported manual-break diagnostics.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_MULTI_COLUMN"), "Explicit break-only final-section column flow should not emit the stale blanket multi-column diagnostic.");

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));
        TestAssert.Equal(1, document.BodyElements.OfType<DocxManualBreakElement>().Count(element => element.Value == "column"));
    }

    public static void DocxReaderPreservesNestedVisibleInlineContainerRuns()
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/nested" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:hyperlink r:id="rIdLink">
                        <w:bookmarkStart w:id="7" w:name="NestedLinkStart"/>
                        <w:ins>
                          <w:r><w:t>Nested</w:t></w:r>
                        </w:ins>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_TRACKED_CHANGES"), "Supported visible inserted runs inside hyperlinks should not be rejected by run-only tracked-change diagnostics.");

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("Nested", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, paragraph.BookmarkAnchors.Count);
        TestAssert.Equal(1, paragraph.Hyperlinks.Count);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].SourceRunCount);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].TextRunCount);
        TestAssert.Equal(6, paragraph.Hyperlinks[0].TextLength);
    }

    public static void DocxReaderPreservesMoveToFinalViewRuns()
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
                      <w:r><w:t>Before </w:t></w:r>
                      <w:moveTo w:id="9" w:author="Author" w:date="2026-06-02T00:00:00Z">
                        <w:r><w:t>Moved</w:t></w:r>
                      </w:moveTo>
                      <w:r><w:t> after</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_TRACKED_CHANGES"), "Final-view moved-to runs should not be rejected as unsupported tracked changes.");

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("Before Moved after", string.Concat(paragraph.Runs.Select(run => run.Text)));
    }

    public static void DocxReaderSplitsNestedVisibleInlineContainerPageBreaks()
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/break" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:hyperlink r:id="rIdLink">
                        <w:ins>
                          <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                        </w:ins>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Text before the nested break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Nested page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Text after the nested break should remain a paragraph fragment.");
        DocxParagraph before = ((DocxParagraphElement)document.BodyElements[0]).Paragraph;
        DocxParagraph after = ((DocxParagraphElement)document.BodyElements[2]).Paragraph;
        TestAssert.Equal("Before", string.Concat(before.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(after.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, before.Hyperlinks.Count);
        TestAssert.Equal(1, after.Hyperlinks.Count);
    }

    public static void DocxReaderSplitsMoveToFinalViewPageBreaks()
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
                      <w:moveTo w:id="9" w:author="Author" w:date="2026-06-02T00:00:00Z">
                        <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                      </w:moveTo>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Moved-to text before the break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Moved-to page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Moved-to text after the break should remain a paragraph fragment.");
        TestAssert.Equal("Before", string.Concat(((DocxParagraphElement)document.BodyElements[0]).Paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(((DocxParagraphElement)document.BodyElements[2]).Paragraph.Runs.Select(run => run.Text)));
    }

    public static void DocxSupportedBodyKeepRulesDoNotEmitUnsupportedKeepDiagnostic()
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
                        <w:keepNext/>
                        <w:keepLines/>
                        <w:widowControl/>
                      </w:pPr>
                      <w:r><w:t>Kept body paragraph</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Following paragraph</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE"), "Body paragraph keep/widow rules are parsed and consumed by page layout, so they should not emit stale unsupported diagnostics.");
    }

    public static void DocxUnsupportedStoryDiagnosticsPreferRelatedPartNames()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
                  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
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
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
                </Relationships>
                """,
            ["word/_rels/comments.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdCommentLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/comment" TargetMode="External"/>
                  <Relationship Id="rIdCommentImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/comment.png"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:commentRangeStart w:id="1"/>
                      <w:r><w:t>Referenced story bodies</w:t></w:r>
                      <w:r><w:commentReference w:id="1"/></w:r>
                      <w:r><w:t>Before</w:t><w:footnoteReference w:id="2" w:customMarkFollows="1"/><w:t>After</w:t></w:r>
                      <w:r><w:endnoteReference w:id="3"/></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:comment w:id="1">
                    <w:p>
                      <w:r><w:t>Comment body</w:t></w:r>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="1">
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:posOffset>228600</wp:posOffset></wp:positionV>
                            <wp:wrapNone/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdCommentImage"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:p><w:hyperlink r:id="rIdCommentLink"><w:r><w:t>Comment link</w:t></w:r></w:hyperlink></w:p>
                    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Comment table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
                  </w:comment>
                </w:comments>
                """,
            ["word/footnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:id="2"><w:p><w:r><w:t>Footnote body</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """,
            ["word/endnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:endnote w:id="3"><w:p><w:r><w:t>Endnote body</w:t></w:r></w:p></w:endnote>
                </w:endnotes>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_COMMENTS" && d.PartName == "/word/comments.xml"), "Comments diagnostics should point to comments.xml when the story body exists.");
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_FOOTNOTE" && d.PartName == "/word/footnotes.xml"), "Footnote diagnostics should point to footnotes.xml when the story body exists.");
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_ENDNOTE" && d.PartName == "/word/endnotes.xml"), "Endnote diagnostics should point to endnotes.xml when the story body exists.");
        TestAssert.True(!diagnostics.Any(d =>
            (d.Id == "DOCX_UNSUPPORTED_COMMENTS" || d.Id == "DOCX_UNSUPPORTED_FOOTNOTE" || d.Id == "DOCX_UNSUPPORTED_ENDNOTE") &&
            d.PartName == "/word/document.xml"), "Story-body diagnostics should not be flattened to document.xml when the related body part exists.");

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream));
        DocxParagraph referenceParagraph = document.Paragraphs.Single();
        TestAssert.Equal(3, referenceParagraph.InlineReferences.Count);
        DocxInlineReference commentReference = referenceParagraph.InlineReferences.Single(reference => reference.Kind == "Comment");
        DocxInlineReference footnoteReference = referenceParagraph.InlineReferences.Single(reference => reference.Kind == "Footnote");
        DocxInlineReference endnoteReference = referenceParagraph.InlineReferences.Single(reference => reference.Kind == "Endnote");
        TestAssert.True(commentReference.Id == "1" && commentReference.CustomMarkFollowsValue is null, "Comment reference markers should be preserved as inline DOCX structure.");
        TestAssert.True(footnoteReference.Id == "2" && footnoteReference.CustomMarkFollowsValue == "1", "Footnote reference markers should preserve custom mark flags.");
        TestAssert.True(endnoteReference.Id == "3" && endnoteReference.CustomMarkFollowsValue is null, "Endnote reference markers should be preserved as inline DOCX structure.");
        TestAssert.Equal(1, commentReference.SourceRunIndex);
        TestAssert.Equal(0, commentReference.RunChildIndex);
        TestAssert.Equal(0, commentReference.TextOffsetInRun);
        TestAssert.Equal(2, footnoteReference.SourceRunIndex);
        TestAssert.Equal(1, footnoteReference.RunChildIndex);
        TestAssert.Equal(6, footnoteReference.TextOffsetInRun);
        TestAssert.Equal(3, endnoteReference.SourceRunIndex);
        TestAssert.Equal(0, endnoteReference.RunChildIndex);
        TestAssert.Equal(0, endnoteReference.TextOffsetInRun);
        TestAssert.Equal(3, document.RelatedStories.Count);
        DocxRelatedStory commentStory = document.RelatedStories.Single(story => story.Kind == "Comment");
        TestAssert.True(commentStory.PartName == "/word/comments.xml" && commentStory.Id == "1" && commentStory.BodyElements.Count == 3 && commentStory.Paragraphs.Count == 2 && commentStory.Tables.Count == 1, "Comment bodies should be preserved as related DOCX stories.");
        TestAssert.Equal(1, commentStory.Paragraphs.Sum(paragraph => paragraph.Hyperlinks.Count));
        TestAssert.Equal("https://example.invalid/comment", commentStory.Paragraphs.SelectMany(paragraph => paragraph.Hyperlinks).Single().Target ?? string.Empty);
        DocxFloatingDrawing commentDrawing = commentStory.FloatingDrawings.Single();
        TestAssert.True(commentDrawing.ImageRelationshipId == "rIdCommentImage" && commentDrawing.SourceParagraphIndex == 0 && commentDrawing.SourceBlockIndex == 0, "Related-story anchored drawings should preserve their owning paragraph, block, and part-local image relationship.");
        TestAssert.True(commentDrawing.HorizontalRelativeFromValue == "page" && commentDrawing.VerticalRelativeFromValue == "paragraph" && commentDrawing.BehindDocumentValue == "1", "Related-story anchored drawing geometry tokens should be preserved structurally before story placement is modeled.");
        TestAssert.True(document.RelatedStories.Any(story => story.Kind == "Footnote" && story.PartName == "/word/footnotes.xml" && story.Id == "2" && story.Paragraphs.Count == 1), "Footnote bodies should be preserved as related DOCX stories.");
        TestAssert.True(document.RelatedStories.Any(story => story.Kind == "Endnote" && story.PartName == "/word/endnotes.xml" && story.Id == "3" && story.Paragraphs.Count == 1), "Endnote bodies should be preserved as related DOCX stories.");

        DocxStructureSnapshot snapshot = new DocxRenderer().InspectStructure(document);
        DocxStructureBlockSnapshot referenceBlock = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal(3, snapshot.InlineReferenceCount);
        TestAssert.Equal(3, snapshot.AnchoredInlineReferenceCount);
        TestAssert.Equal(3, snapshot.ResolvedInlineReferenceCount);
        TestAssert.Equal(6, snapshot.MaxInlineReferenceTextOffsetInRun);
        TestAssert.Equal(3, referenceBlock.InlineReferenceCount);
        TestAssert.Equal(3, referenceBlock.AnchoredInlineReferenceCount);
        TestAssert.Equal(3, referenceBlock.ResolvedInlineReferenceCount);
        TestAssert.Equal(6, referenceBlock.MaxInlineReferenceTextOffsetInRun);
        TestAssert.Equal(1, referenceBlock.CommentReferenceCount);
        TestAssert.Equal(1, referenceBlock.FootnoteReferenceCount);
        TestAssert.Equal(1, referenceBlock.EndnoteReferenceCount);
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Body" && story.InlineReferenceCount == 3 && story.ResolvedInlineReferenceCount == 3), "Structure snapshots should expose body inline story-reference ownership.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Comment" && story.Scope == "/word/comments.xml" && story.VariantType == "1" && story.BlockCount == 3 && story.ParagraphCount == 2 && story.TableCount == 1 && story.TextLength == 37 && story.HyperlinkCount == 1 && story.ExternalHyperlinkCount == 1 && story.FloatingDrawingCount == 1), "Structure snapshots should expose comment story ownership, hyperlinks, anchored drawings, and table metrics.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Footnote" && story.Scope == "/word/footnotes.xml" && story.VariantType == "2" && story.TextLength == 13), "Structure snapshots should expose footnote story text metrics.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Endnote" && story.Scope == "/word/endnotes.xml" && story.VariantType == "3" && story.TextLength == 12), "Structure snapshots should expose endnote story text metrics.");
        TestAssert.Equal(3, snapshot.InlineReferences.Count);
        DocxStructureInlineReferenceSnapshot commentReferenceSnapshot = snapshot.InlineReferences.Single(reference => reference.Kind == "Comment");
        TestAssert.True(commentReferenceSnapshot.SourceBlockIndex == 0 && commentReferenceSnapshot.SourceBlockKind == "Paragraph" && commentReferenceSnapshot.SourceRunIndex == 1 && commentReferenceSnapshot.TextOffsetInRun == 0, "Comment reference snapshot should preserve the source marker coordinates.");
        TestAssert.True(commentReferenceSnapshot.ResolvedStoryKind == "Comment" && commentReferenceSnapshot.ResolvedStoryPartName == "/word/comments.xml" && commentReferenceSnapshot.ResolvedStoryId == "1" && commentReferenceSnapshot.ResolvedStoryBlockCount == 3 && commentReferenceSnapshot.ResolvedStoryTextLength == 37, "Comment reference snapshot should resolve to the comment story body.");
        DocxStructureInlineReferenceSnapshot footnoteReferenceSnapshot = snapshot.InlineReferences.Single(reference => reference.Kind == "Footnote");
        TestAssert.True(footnoteReferenceSnapshot.CustomMarkFollowsValue == "1" && footnoteReferenceSnapshot.SourceRunIndex == 2 && footnoteReferenceSnapshot.TextOffsetInRun == 6, "Footnote reference snapshot should preserve custom mark and source marker offsets.");
        TestAssert.True(footnoteReferenceSnapshot.ResolvedStoryKind == "Footnote" && footnoteReferenceSnapshot.ResolvedStoryPartName == "/word/footnotes.xml" && footnoteReferenceSnapshot.ResolvedStoryId == "2" && footnoteReferenceSnapshot.ResolvedStoryTextLength == 13, "Footnote reference snapshot should resolve to the footnote story body.");
        DocxStructureInlineReferenceSnapshot endnoteReferenceSnapshot = snapshot.InlineReferences.Single(reference => reference.Kind == "Endnote");
        TestAssert.True(endnoteReferenceSnapshot.SourceRunIndex == 3 && endnoteReferenceSnapshot.TextOffsetInRun == 0, "Endnote reference snapshot should preserve source marker offsets.");
        TestAssert.True(endnoteReferenceSnapshot.ResolvedStoryKind == "Endnote" && endnoteReferenceSnapshot.ResolvedStoryPartName == "/word/endnotes.xml" && endnoteReferenceSnapshot.ResolvedStoryId == "3" && endnoteReferenceSnapshot.ResolvedStoryTextLength == 12, "Endnote reference snapshot should resolve to the endnote story body.");

        DocxLayoutSnapshot layoutSnapshot = new DocxRenderer().InspectLayout(document);
        TestAssert.Equal(3, layoutSnapshot.RelatedStories.Count);
        DocxRelatedStoryLayoutSnapshot commentLayout = layoutSnapshot.RelatedStories.Single(story => story.Kind == "Comment");
        TestAssert.True(commentLayout.PartName == "/word/comments.xml" && commentLayout.Id == "1" && commentLayout.BlockCount == 3 && commentLayout.ParagraphCount == 2 && commentLayout.TableCount == 1, "Related-story layout snapshots should preserve comment story ownership without flattening it into body layout.");
        TestAssert.True(commentLayout.TextLineCount >= 2 && commentLayout.TableCellTextLineCount >= 1 && commentLayout.TableRowCount == 1 && commentLayout.FloatingDrawingCount == 1 && commentLayout.TextLength == 37 && commentLayout.ContentHeight > 0d, "Comment story layout should measure paragraph text and table rows while preserving unpaged anchored drawing ownership.");
        TestAssert.True(commentLayout.Items.Count(item => item.Kind == "TextLine") >= 2 && commentLayout.Items.Count(item => item.Kind == "TableRow") == 1 && commentLayout.TableRows.Count == 1, "Related-story snapshots should expose private-safe item and table-row ownership for future story placement.");
        TestAssert.True(commentLayout.SourceBlocks.Count == 3 && commentLayout.SourceBlocks.Any(block => block.Kind == "Table" && block.TableRowCount == 1) && commentLayout.SourceBlocks.Count(block => block.Kind == "Paragraph") == 2, "Related-story snapshots should expose private-safe source-block summaries without assigning fake page indexes.");
        TestAssert.True(layoutSnapshot.RelatedStories.Any(story => story.Kind == "Footnote" && story.PartName == "/word/footnotes.xml" && story.Id == "2" && story.TextLineCount >= 1 && story.TableRowCount == 0 && story.ContentHeight > 0d), "Footnote story layout should be measured as related-story content.");
        TestAssert.True(layoutSnapshot.RelatedStories.Any(story => story.Kind == "Endnote" && story.PartName == "/word/endnotes.xml" && story.Id == "3" && story.TextLineCount >= 1 && story.TableRowCount == 0 && story.ContentHeight > 0d), "Endnote story layout should be measured as related-story content.");

        DocxFontPlan fontPlan = DocxFontPlan.Create(document, new MapFontResolver([], "Fallback"));
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Comment body"), "Related story runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Comment link"), "Related story hyperlink runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Comment table"), "Related story table runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Footnote body"), "Footnote runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Endnote body"), "Endnote runs should participate in DOCX font planning.");
    }

    public static void DocxRelatedStoryLayoutOwnsInlineImages()
    {
        DocxTextRun bodyRun = new("Body", 12d, "000000", Bold: false, Italic: false, Underline: false, UnderlineValue: null, FontFamily: null);
        DocxParagraph bodyParagraph = new(
            [bodyRun],
            [],
            StyleId: null,
            DocxTextAlignment.Left,
            AlignmentValue: null,
            SpacingBeforePoints: 0d,
            SpacingAfterPoints: 0d,
            LineSpacingFactor: 1d,
            LineSpacingPoints: null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            ListLabel: null);
        DocxInlineImage storyImage = new(48d, 24d, "image/png", [1, 2, 3], "/word/media/comment.png");
        DocxParagraph imageParagraph = new(
            [],
            [storyImage],
            StyleId: null,
            DocxTextAlignment.Center,
            AlignmentValue: "center",
            SpacingBeforePoints: 0d,
            SpacingAfterPoints: 0d,
            LineSpacingFactor: 1d,
            LineSpacingPoints: null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            ListLabel: null);
        DocxRelatedStory commentStory = new(
            "Comment",
            "/word/comments.xml",
            "9",
            [new DocxParagraphElement(imageParagraph)],
            [],
            []);
        DocxDocument document = new DocxDocument(612d, 792d)
        {
            BodyElements = [new DocxParagraphElement(bodyParagraph)],
            Paragraphs = [bodyParagraph],
            RelatedStories = [commentStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine().Create(document, new FamilyWidthTextMeasurer()));
        DocxRelatedStoryLayoutSnapshot storySnapshot = snapshot.RelatedStories.Single();

        TestAssert.True(storySnapshot.Kind == "Comment" && storySnapshot.PartName == "/word/comments.xml" && storySnapshot.Id == "9", "Related-story layout snapshots should preserve the owning story identity.");
        TestAssert.True(storySnapshot.BlockCount == 1 && storySnapshot.ParagraphCount == 1 && storySnapshot.TableCount == 0, "Related-story layout snapshots should derive counts from the body block stream, not stale parallel inventories.");
        TestAssert.True(storySnapshot.TextLineCount == 0 && storySnapshot.InlineImageCount == 1 && storySnapshot.ContentHeight >= 24d, "Related-story inline images should be owned by story layout instead of only contributing anonymous height.");
        DocxLayoutItemSnapshot imageItem = storySnapshot.Items.Single(item => item.Kind == "InlineImage");
        TestAssert.True(imageItem.SourceBlockIndex == 0 && imageItem.SourceParagraphIndex == 0 && imageItem.Width == 48d && imageItem.Height == 24d, "Related-story image item snapshots should carry private-safe source coordinates and geometry.");
        DocxRelatedStorySourceBlockSnapshot imageBlock = storySnapshot.SourceBlocks.Single();
        TestAssert.True(imageBlock.Kind == "InlineImage" && imageBlock.InlineImageCount == 1 && imageBlock.ItemCount == 1 && imageBlock.ConsumedHeight >= 24d, "Related-story source-block snapshots should summarize image-only story blocks without page ownership.");
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
                      <w:pPr><w:ind w:left="720" w:firstLine="120"/></w:pPr>
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

    public static void DocxUnsupportedTableBorderStylesEmitDiagnostics()
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="StyledTable"/>
                        <w:tblBorders>
                          <w:top w:val="single" w:sz="8"/>
                          <w:insideH w:val="nil"/>
                          <w:insideV w:val="none"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tr><w:tc><w:p><w:r><w:t>Styled</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblPr>
                      <w:tblBorders>
                        <w:bottom w:val="double" w:sz="12"/>
                      </w:tblBorders>
                    </w:tblPr>
                  </w:style>
                </w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(
            diagnostics.Any(d => d.Id == "DOCX_TABLE_BORDER_STYLE" && d.PartName == "/word/styles.xml" && d.Fallback == "Approximated"),
            "Unsupported DOCX table border styles should emit a style-part diagnostic instead of being silently flattened to solid borders.");
    }

    public static void DocxSupportedTableStyleAtomsDoNotEmitBroadTableStyleDiagnostic()
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="StyledTable"/></w:tblPr>
                      <w:tr><w:tc><w:p><w:r><w:t>Styled</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblPr>
                      <w:tblW w:w="5000" w:type="pct"/>
                      <w:tblInd w:w="120" w:type="dxa"/>
                      <w:tblCellSpacing w:w="20" w:type="dxa"/>
                    </w:tblPr>
                    <w:tblStylePr w:type="firstRow">
                      <w:tcPr><w:shd w:fill="DDDDDD"/></w:tcPr>
                      <w:pPr><w:jc w:val="center"/></w:pPr>
                      <w:rPr><w:b/><w:i/><w:caps/><w:color w:val="336699"/><w:sz w:val="22"/></w:rPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_STYLE_TABLE_STYLE", ids);
        TestAssert.DoesNotContain("DOCX_STYLE_TABLE_COMPLEX_SCRIPT_RUN", ids);
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_TABLE_STYLE", ids);
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

    public static void DocxSupportedStyleKeepRulesDoNotEmitDiagnostics()
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:pStyle w:val="SupportedKeep"/></w:pPr><w:r><w:t>One</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pStyle w:val="SupportedKeep"/></w:pPr><w:r><w:t>Two</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="SupportedKeep">
                    <w:pPr>
                      <w:keepNext/>
                      <w:keepLines/>
                    </w:pPr>
                  </w:style>
                </w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_STYLE_PARAGRAPH_KEEP_RULE", ids);
    }

    public static void DocxSupportedStyleSpacingVariantsDoNotEmitDiagnostics()
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:pStyle w:val="SupportedSpacing"/></w:pPr><w:r><w:t>One</w:t></w:r></w:p>
                    <w:p><w:pPr><w:pStyle w:val="SupportedSpacing"/></w:pPr><w:r><w:t>Two</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="SupportedSpacing">
                    <w:pPr>
                      <w:contextualSpacing/>
                      <w:spacing w:beforeLines="120" w:afterLines="240"/>
                    </w:pPr>
                  </w:style>
                </w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_STYLE_PARAGRAPH_SPACING", ids);
    }

    private static DocxDocument CreateFontPlanDocument(DocxTextRun run, DocxFontCatalog fontCatalog)
    {
        return CreateFontPlanDocument([run], fontCatalog);
    }

    private static DocxDocument CreateFontPlanDocument(IReadOnlyList<DocxTextRun> runs, DocxFontCatalog fontCatalog)
    {
        return CreateFontPlanDocument([CreateFontPlanParagraph(runs)], fontCatalog);
    }

    private static DocxParagraph CreateFontPlanParagraph(DocxTextRun run)
    {
        return CreateFontPlanParagraph([run]);
    }

    private static DocxParagraph CreateFontPlanParagraph(IReadOnlyList<DocxTextRun> runs)
    {
        return new DocxParagraph(
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
    }

    private static DocxDocument CreateFontPlanDocument(IReadOnlyList<DocxParagraph> paragraphs, DocxFontCatalog fontCatalog)
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
            paragraphs.Select(paragraph => new DocxParagraphElement(paragraph)).ToArray(),
            paragraphs,
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

    private sealed class FamilyWidthTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            double width = run?.FontFamily == "Wide" ? 40d : 5d;
            return text.Length * width;
        }

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
        {
            return fontSize;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            return run?.FontFamily == "Label Metrics" ? fontSize * 1.1d : fontSize;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
        {
            return run?.FontFamily == "Label Metrics" ? fontSize * 0.3d : fontSize * 0.2d;
        }
    }

    private sealed class FractionalLineHeightTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            return text.Length * 5d;
        }

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
        {
            return 11.98875d;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            return fontSize;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
        {
            return fontSize * 0.2d;
        }
    }

    private sealed class FontSizeWidthTextMeasurer : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            return text.Length * fontSize;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            return fontSize;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
        {
            return fontSize * 0.2d;
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
        DocxParagraphKeepRules? keepRules = null,
        DocxParagraphIndent? indent = null)
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
            null)
        {
            Indent = indent ?? DocxParagraphIndent.Empty
        };
    }

    private static double[] ExtractTextBaselines(string pdf)
    {
        return Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .Distinct()
            .ToArray();
    }

    private static double[] ExtractTextBaselinesAtX(string pdf, double x)
    {
        string escapedX = Regex.Escape(x.ToString("0.###", CultureInfo.InvariantCulture));
        return Regex.Matches(pdf, $@"1 0 0 1 {escapedX}(?:\.\d+)? (?<y>-?\d+(?:\.\d+)?) Tm")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .Distinct()
            .ToArray();
    }

    private static int CountPdfTextShows(string pdf)
    {
        return CountOccurrences(pdf, "> Tj") + CountOccurrences(pdf, "] TJ");
    }

    private static int FirstPdfTextShowIndex(string pdf)
    {
        int tj = pdf.IndexOf("> Tj", StringComparison.Ordinal);
        int positioned = pdf.IndexOf("] TJ", StringComparison.Ordinal);
        return MinNonNegative(tj, positioned);
    }

    private static int LastPdfTextShowIndex(string pdf)
    {
        return Math.Max(
            pdf.LastIndexOf("> Tj", StringComparison.Ordinal),
            pdf.LastIndexOf("] TJ", StringComparison.Ordinal));
    }

    private static int CountOccurrences(string text, string value)
    {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static int MinNonNegative(int first, int second)
    {
        if (first < 0)
        {
            return second;
        }

        return second < 0 ? first : Math.Min(first, second);
    }
}
