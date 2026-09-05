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

internal static class DocxTextTests
{
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 1, "Expected DOCX paragraph text to render as a PDF text-show operation.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
                      <w:wordWrap w:val="0"/>
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("Risky", paragraph.StyleId ?? string.Empty);
        TestAssert.Equal(18d, paragraph.SpacingBeforePoints);
        TestAssert.Equal(14d, paragraph.SpacingAfterPoints);
        TestAssert.Equal("360", paragraph.Spacing.BeforeValue ?? string.Empty);
        TestAssert.Equal(string.Empty, paragraph.Spacing.AfterValue ?? string.Empty);
        TestAssert.Equal(string.Empty, paragraph.Spacing.AfterLinesValue ?? string.Empty);
        TestAssert.Equal(string.Empty, paragraph.Spacing.BeforeAutoSpacingValue ?? string.Empty);
        TestAssert.Equal("1", paragraph.Spacing.AfterAutoSpacingValue ?? string.Empty);
        DocxStructureBlockSnapshot block = DocxStructureSnapshot.FromDocument(document).Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal("1", block.AfterAutoSpacingValue ?? string.Empty);
        TestAssert.Equal(string.Empty, block.BeforeAutoSpacingValue ?? string.Empty);
        TestAssert.True(block.WordWrap == false, "Structure snapshots should expose paragraph wordWrap without document text.");
        TestAssert.Equal("0", block.WordWrapValue ?? string.Empty);
        TestAssert.Equal("480", paragraph.Spacing.LineValue ?? string.Empty);
        TestAssert.Equal("exact", paragraph.Spacing.LineRuleValue ?? string.Empty);
        DocxStructureStyleUsageSnapshot styleUsage = DocxStructureSnapshot.FromDocument(document).StyleUsages.Single(usage => usage.Kind == "Paragraph");
        TestAssert.Equal(1, styleUsage.BeforeSpacingTokenParagraphCount);
        TestAssert.Equal(1, styleUsage.AfterSpacingTokenParagraphCount);
        TestAssert.Equal(0, styleUsage.BeforeAutoSpacingParagraphCount);
        TestAssert.Equal(1, styleUsage.AfterAutoSpacingParagraphCount);
        TestAssert.Equal(0, styleUsage.BeforeLinesSpacingParagraphCount);
        TestAssert.Equal(0, styleUsage.AfterLinesSpacingParagraphCount);
        TestAssert.Equal(1, styleUsage.ContextualSpacingParagraphCount);
        TestAssert.Equal(1, styleUsage.ExactLineSpacingParagraphCount);
        TestAssert.Equal(0, styleUsage.AtLeastLineSpacingParagraphCount);
        TestAssert.Equal(0, styleUsage.AutoLineSpacingParagraphCount);
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
        TestAssert.True(paragraph.WordWrap == false, "Style wordWrap off should survive the paragraph cascade.");
        TestAssert.Equal("0", paragraph.WordWrapValue ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph paragraph = document.Paragraphs.Single();
        DocxTextRun run = paragraph.Runs.Single();
        TestAssert.Equal("Child", paragraph.StyleId ?? string.Empty);
        TestAssert.True(paragraph.StyleResolution.StyleFound, "Paragraph style resolution should record that the referenced style was found.");
        TestAssert.Equal(2, paragraph.StyleResolution.StyleDepth);
        TestAssert.True(paragraph.StyleResolution.HasDirectParagraphProperties == false, "A pPr containing only pStyle should not be treated as a direct layout override.");
        TestAssert.True(paragraph.StyleResolution.HasDocumentDefaultParagraphProperties == false, "This package has no paragraph defaults.");
        TestAssert.True(paragraph.StyleResolution.HasTableStyleParagraphProperties == false, "Body paragraphs should not report table-style paragraph properties.");
        TestAssert.Equal(6d, paragraph.SpacingBeforePoints);
        TestAssert.Equal(0d, paragraph.SpacingAfterPoints);
        TestAssert.Equal("120", paragraph.Spacing.BeforeValue ?? string.Empty);
        TestAssert.Equal("0", paragraph.Spacing.AfterValue ?? string.Empty);
        TestAssert.Equal("360", paragraph.Spacing.LineValue ?? string.Empty);
        TestAssert.Equal("auto", paragraph.Spacing.LineRuleValue ?? string.Empty);
        TestAssert.Equal(1.5d, paragraph.LineSpacingFactor);
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        TestAssert.Equal("Child", effective.StyleId ?? string.Empty);
        TestAssert.Equal(6d, effective.SpacingBeforePoints);
        TestAssert.Equal(0d, effective.SpacingAfterPoints);
        TestAssert.Equal(1.5d, effective.LineSpacingFactor);
        TestAssert.True(effective.StyleResolution.StyleFound, "The layout-ready effective paragraph model should preserve style provenance.");
        TestAssert.True(effective.KeepRules.KeepNext != true, "The effective model should carry resolved keep rules without inventing them.");
        TestAssert.Equal(14d, run.FontSize);
        TestAssert.Equal("112233", run.ColorHex ?? string.Empty);
        TestAssert.True(run.Bold, "Child paragraph style run properties should merge over the inherited base run size/color.");

        DocxStyleDefinitionSummary childStyle = document.StyleCatalog.ParagraphStyles.Single(style => style.StyleId == "Child");
        TestAssert.True(childStyle.BasedOnStyleId == "Base" && childStyle.HasParagraphProperties && childStyle.HasRunProperties, "The DOCX style catalog should retain private-safe paragraph style topology after resolved properties are applied.");
        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(2, structure.StyleCatalog.ParagraphStyles.Count);
        DocxStructureBlockSnapshot block = structure.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.True(block.ParagraphStyleFound == true, "Structure snapshots should expose private-safe paragraph style resolution.");
        TestAssert.Equal(2, block.ParagraphStyleDepth ?? 0);
        TestAssert.True(block.HasDirectParagraphProperties == false, "Structure snapshots should distinguish pStyle-only pPr from direct paragraph overrides.");

        DocxLayoutItemSnapshot line = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document).Pages
            .SelectMany(page => page.Items)
            .Single(item => item.Kind == "TextLine");
        TestAssert.Equal("Child", line.ParagraphStyleId ?? string.Empty);
        TestAssert.True(line.ParagraphStyleFound == true, "Layout snapshots should carry paragraph style resolution for placed lines.");
        TestAssert.Equal(2, line.ParagraphStyleDepth ?? 0);
        TestAssert.True(line.HasDirectParagraphProperties == false, "Placed line snapshots should preserve pStyle-only provenance.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraphElement[] paragraphs = document.BodyElements.OfType<DocxParagraphElement>().ToArray();
        TestAssert.Equal(3, paragraphs.Length);
        TestAssert.Equal(1, paragraphs[1].Paragraph.Runs.Count);
        TestAssert.Equal(string.Empty, paragraphs[1].Paragraph.Runs[0].Text);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal(36d, paragraph.SpacingBeforePoints);
        TestAssert.Equal(48d, paragraph.SpacingAfterPoints);
        TestAssert.Equal("150", paragraph.Spacing.BeforeLinesValue ?? string.Empty);
        TestAssert.Equal("200", paragraph.Spacing.AfterLinesValue ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextRun run = document.Paragraphs.Single().Runs.Single();
        TestAssert.Equal("Paragraph Sans", run.FontFamily ?? string.Empty);
        TestAssert.Equal("Paragraph Sans", run.Fonts.Ascii ?? string.Empty);
        TestAssert.Equal("Character Sans", run.Fonts.HighAnsi ?? string.Empty);
        TestAssert.Equal("Paragraph East", run.Fonts.EastAsia ?? string.Empty);
        TestAssert.Equal("majorHAnsi", run.Fonts.AsciiTheme ?? string.Empty);
        TestAssert.Equal("minorHAnsi", run.Fonts.HighAnsiTheme ?? string.Empty);
        TestAssert.Equal("majorBidi", run.Fonts.ComplexScriptTheme ?? string.Empty);
        TestAssert.Equal("Emphasis", run.StyleResolution.CharacterStyleId ?? string.Empty);
        TestAssert.True(run.StyleResolution.CharacterStyleFound, "Run style provenance should record resolved character styles.");
        TestAssert.Equal(1, run.StyleResolution.CharacterStyleDepth);
        TestAssert.True(run.StyleResolution.HasDocumentDefaultRunProperties, "Run style provenance should retain document-default participation.");
        TestAssert.True(run.StyleResolution.HasParagraphStyleRunProperties, "Run style provenance should retain paragraph-style run contribution.");
        TestAssert.True(run.StyleResolution.HasCharacterStyleRunProperties, "Run style provenance should retain character-style run contribution.");
        TestAssert.True(run.StyleResolution.HasDirectRunProperties, "Direct rPr beyond rStyle should remain distinguishable from style references.");
        TestAssert.True(run.StyleResolution.HasTableStyleRunProperties == false, "Body paragraph runs should not report table-style run contribution.");
        TestAssert.Equal("Emphasis", run.EffectiveProperties.StyleResolution.CharacterStyleId ?? string.Empty);

        DocxLayoutItemSnapshot line = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document).Pages
            .SelectMany(page => page.Items)
            .Single(item => item.Kind == "TextLine");
        TestAssert.True(line.CharacterStyleTextSegmentCount > 0, "Layout snapshots should expose character-style text segment provenance.");
        TestAssert.True(line.DirectRunPropertyTextSegmentCount > 0, "Layout snapshots should expose direct run-property text segment provenance.");
        TestAssert.True(line.ParagraphStyleRunPropertyTextSegmentCount > 0, "Layout snapshots should expose paragraph-style run-property text segment provenance.");
        TestAssert.True(line.DocumentDefaultRunPropertyTextSegmentCount > 0, "Layout snapshots should expose document-default run-property text segment provenance.");
        TestAssert.Equal(0, line.TableStyleRunPropertyTextSegmentCount ?? -1);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextRun run = document.Paragraphs.Single().Runs.Single();
        TestAssert.Equal("Base Sans", run.FontFamily ?? string.Empty);
        TestAssert.Equal("Base Sans", run.Fonts.Ascii ?? string.Empty);
        TestAssert.Equal("Base Sans", run.Fonts.HighAnsi ?? string.Empty);
        TestAssert.Equal("Child East", run.Fonts.EastAsia ?? string.Empty);
    }

    public static void DocxFontPlanResolvesImplicitDefaultTypefaceWithRunStyle()
    {
        var run = new DocxTextRun("Styled default", 11d, null, true, true, false, null, null);
        DocxDocument document = DocxTests.CreateFontPlanDocument(run, DocxFontCatalog.Empty);
        var resolver = new MapFontResolver([DocxRenderer.DefaultDocumentTypefaceRequest], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Single();

        TestAssert.Equal(DocxTypefaceResolutionSource.ResolverFallback, resolved.Source);
        TestAssert.Equal(DocxRenderer.DefaultDocumentTypefaceRequest, resolved.RequestedFamily ?? string.Empty);
        TestAssert.True(resolved.Resolution?.Bold == true && resolved.Resolution?.Italic == true, "Implicit default runs should preserve requested bold and italic in the fallback font request.");
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FontSizeWidthTextMeasurer(), CancellationToken.None)
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
        var superscriptRun = new DocxTextRun("2", 16d, null, false, false, false, null, "Body", 0d, false, "superscript", false, null, false, null, null, null, null, null, false, null, false, null, null);
        var subscriptRun = new DocxTextRun("n", 16d, null, false, false, false, null, "Body", 0d, false, "subscript", false, null, false, null, null, null, null, null, false, null, false, null, null);
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FontSizeWidthTextMeasurer(), CancellationToken.None)
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FontSizeWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal("A B", lines[0].Text);
        TestAssert.Equal(40d, lines[0].Width);
        TestAssert.Equal(2, lines[0].Segments.Count);
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
                    <w:rPr><w:u w:val="wave" w:color="00AA00"/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:u w:val="single" w:color="0000FF"/></w:rPr><w:t>Single</w:t></w:r>
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxTextRun[] runs = document.Paragraphs[0].Runs.ToArray();

        TestAssert.True(runs[0].Underline, "Expected w:u single to keep underline enabled.");
        TestAssert.Equal("single", runs[0].UnderlineValue ?? string.Empty);
        TestAssert.Equal("0000FF", runs[0].UnderlineColorHex ?? string.Empty);
        TestAssert.True(!runs[1].Underline, "Expected w:u none to disable underline.");
        TestAssert.Equal("none", runs[1].UnderlineValue ?? string.Empty);
        TestAssert.True(runs[2].Underline, "Expected inherited w:u wave to keep underline enabled.");
        TestAssert.Equal("wave", runs[2].UnderlineValue ?? string.Empty);
        TestAssert.Equal("00AA00", runs[2].UnderlineColorHex ?? string.Empty);
        TestAssert.True(runs[3].UnderlineValue is null, "Expected missing underline to keep a null source token.");
        TestAssert.True(runs[3].UnderlineColorHex is null, "Expected missing underline color to keep a null source token.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="double"/></w:rPr><w:t> DoubleUnder</w:t></w:r>
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
        TestAssert.True(filledRectangles >= 6, $"Expected underline, double-underline, strike, and double-strike to render as filled metric rectangles; found {filledRectangles}.");
    }

    public static void DocxParagraphRendererUsesUnderlineColorToken()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:color w:val="FF0000"/><w:u w:val="single" w:color="0000FF"/></w:rPr><w:t>Under</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int underlineColorIndex = pdf.IndexOf("0 0 1 rg", StringComparison.Ordinal);
        int underlineRectangleIndex = underlineColorIndex < 0 ? -1 : pdf.IndexOf(" re f", underlineColorIndex, StringComparison.Ordinal);
        TestAssert.True(underlineColorIndex >= 0 && underlineRectangleIndex > underlineColorIndex, "Underline decorations should use the w:u color token instead of the run text color.");
    }

    public static void DocxParagraphRendererDrawsWaveUnderlineAsStrokedSegments()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="wave"/></w:rPr><w:t>Wavy underline</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(DocxTests.CountOccurrences(pdf, " l S") >= 2, "Expected w:u wave to render as repeated stroked underline segments instead of a single filled rectangle.");
    }

    public static void DocxParagraphRendererDrawsSegmentedUnderlineStylesAsFragments()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="dash"/></w:rPr><w:t>Dash style</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="dotted"/></w:rPr><w:t> Dotted style</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="dotDash"/></w:rPr><w:t> Dot dash style</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="dotDotDash"/></w:rPr><w:t> Dot dot dash style</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int filledRectangles = DocxTests.CountOccurrences(pdf, " re f");
        TestAssert.True(filledRectangles >= 12, $"Expected segmented underline styles to render as repeated filled fragments; found {filledRectangles}.");
    }

    public static void DocxParagraphRendererDrawsThickUnderlineWithHeavierStroke()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t>Single</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="thick"/></w:rPr><w:t> Thick</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] heights = Regex.Matches(pdf, @"(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) (?<width>\d+(?:\.\d+)?) (?<height>\d+(?:\.\d+)?) re f")
            .Select(match => double.Parse(match.Groups["height"].Value, CultureInfo.InvariantCulture))
            .ToArray();
        TestAssert.True(heights.Length >= 2, "Expected single and thick underlines to emit filled rectangles.");
        TestAssert.True(heights.Max() > heights.Min() * 1.2d, "Expected w:u thick to emit a visibly heavier underline rectangle than w:u single.");
    }

    public static void DocxParagraphRendererDrawsOfficialHeavyDashUnderlineTokensAsFragments()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="dashDotHeavy"/></w:rPr><w:t>Dash dot heavy</w:t></w:r>
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="dashDotDotHeavy"/></w:rPr><w:t> Dash dot dot heavy</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int filledRectangles = DocxTests.CountOccurrences(pdf, " re f");
        TestAssert.True(filledRectangles >= 6, $"Expected official heavy dash underline tokens to render as repeated filled fragments; found {filledRectangles}.");
    }

    public static void DocxParagraphRendererDrawsWordsUnderlineOnlyUnderWords()
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
                      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:u w:val="words"/></w:rPr><w:t>Alpha Beta</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int filledRectangles = DocxTests.CountOccurrences(pdf, " re f");
        TestAssert.True(filledRectangles >= 2, $"Expected w:u words to emit separate underline rectangles around spaces; found {filledRectangles}.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxTextRun[] runs = document.Paragraphs[0].Runs.ToArray();

        TestAssert.True(runs[1].Hidden, "Expected inherited w:vanish to mark the run as hidden.");
        TestAssert.True(runs[1].HiddenValue is null, "Expected val-less hidden style to keep a null source token.");
        TestAssert.True(!runs[2].Hidden, "Expected w:vanish val=0 to keep the run visible.");
        TestAssert.Equal("0", runs[2].HiddenValue ?? string.Empty);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxTextRun[] runs = document.Paragraphs[0].Runs.ToArray();

        TestAssert.Equal("Before ", runs[0].Text);
        TestAssert.Equal("2026", runs[1].Text);
        TestAssert.Equal("336699", runs[1].ColorHex ?? string.Empty);
        TestAssert.Equal(" After", runs[2].Text);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal(1, snapshot.HyperlinkCount);
        TestAssert.Equal(1, snapshot.ExternalHyperlinkCount);
        TestAssert.Equal(0, snapshot.InternalHyperlinkCount);
        TestAssert.Equal(1, block.HyperlinkCount);
        TestAssert.Equal(1, block.ExternalHyperlinkCount);
        TestAssert.Equal(0, block.InternalHyperlinkCount);
    }
}
