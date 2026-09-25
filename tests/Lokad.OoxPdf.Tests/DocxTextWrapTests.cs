using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTextWrapTests
{
    public static void DocxParagraphLayoutPreservesSoftLineBreaks()
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "AlphaBeta".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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

    public static void DocxParagraphLayoutUsesAuthoredRightTabStops()
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
                      <w:pPr><w:tabs><w:tab w:val="right" w:pos="1440"/></w:tabs></w:pPr>
                      <w:r><w:t>A</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>BB</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("A", line.Segments[0].Text);
        TestAssert.Equal("BB", line.Segments[1].Text);
        TestAssert.Equal(62d, line.Segments[1].X - line.X);
        TestAssert.Equal(72d, line.Width);
    }

    public static void DocxParagraphLayoutUsesAuthoredCenterTabStops()
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
                      <w:pPr><w:tabs><w:tab w:val="center" w:pos="1440"/></w:tabs></w:pPr>
                      <w:r><w:t>A</w:t><w:tab/><w:t>BB</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("A", line.Segments[0].Text);
        TestAssert.Equal("BB", line.Segments[1].Text);
        TestAssert.Equal(67d, line.Segments[1].X - line.X);
        TestAssert.Equal(77d, line.Width);
    }

    public static void DocxParagraphLayoutUsesAuthoredDecimalTabStops()
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
                      <w:pPr><w:tabs><w:tab w:val="decimal" w:pos="1440"/></w:tabs></w:pPr>
                      <w:r><w:t>A</w:t><w:tab/><w:t>12.3</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("A", line.Segments[0].Text);
        TestAssert.Equal("12.3", line.Segments[1].Text);
        TestAssert.Equal(62d, line.Segments[1].X - line.X);
        TestAssert.Equal(82d, line.Width);
    }

    public static void DocxParagraphLayoutRightAlignsDecimalTabsWithoutDecimalSeparator()
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
                      <w:pPr><w:tabs><w:tab w:val="decimal" w:pos="1440"/></w:tabs></w:pPr>
                      <w:r><w:t>A</w:t><w:tab/><w:t>123</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal("123", line.Segments[1].Text);
        TestAssert.Equal(57d, line.Segments[1].X - line.X);
        TestAssert.Equal(72d, line.Width);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("Alpha\u00A0Beta ", lines[0].Text);
        TestAssert.Equal("Gamma", lines[1].Text);
    }

    public static void DocxParagraphLayoutBreaksOverwideSoftHyphenatedTokens()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABC\u00ADDEFG", 10d, null, false, false, false, null, null)],
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
            45d,
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("ABC\u00AD", lines[0].Text);
        TestAssert.Equal("DEFG", lines[1].Text);
        TestAssert.True(lines[0].EndsWithIntraTokenBreak, "Soft-hyphen body splits should be marked as intra-token line endings.");
    }
    public static void DocxParagraphLayoutBreaksOverlongTokenAfterHyphenWhenPrefixFits()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("AA BB-well C", 10d, null, false, false, false, null, null)],
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
            50d,
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("AA BB-", lines[0].Text);
        TestAssert.Equal("well C", lines[1].Text);
        TestAssert.True(lines[0].EndsWithIntraTokenBreak, "Greedy hyphen splits should be marked as intra-token line endings.");
    }


    public static void DocxParagraphLayoutSuppressesUnbrokenSoftHyphens()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABC\u00ADDEFG", 10d, null, false, false, false, null, null)],
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
            67d,
            200d,
            15d,
            15d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [],
            []);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal("ABCDEFG", line.Text);
        TestAssert.Equal("ABCDEFG", string.Concat(line.Segments.Select(segment => segment.Text)));
        TestAssert.True(!line.EndsWithIntraTokenBreak, "A discretionary soft hyphen should stay hidden when the token fits without breaking.");
        TestAssert.True(Math.Abs(line.Width - 35d) < 0.001d, "Unbroken soft hyphens should not consume layout width.");
    }

    public static void DocxParagraphLayoutPreservesSourceOffsetsAfterHiddenSoftHyphens()
    {
        var run = new DocxTextRun("AB\u00ADCD", 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 7,
            SourceTextOffsetInRun = 0
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

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal("ABCD", line.Text);
        TestAssert.True(
            line.Segments.Any(segment => segment.SourceTextRunIndex == 7 && segment.SourceTextOffsetInRun == 3 && segment.Text == "CD"),
            "Hidden soft hyphens should not shift source offsets for later rendered text in the same run.");
    }

    public static void DocxParagraphLayoutBreaksAtHiddenZeroWidthSpaces()
    {
        var run = new DocxTextRun("ABC\u200BDEFG", 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 8,
            SourceTextOffsetInRun = 0
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
            null);
        var document = new DocxDocument(
            40d,
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("ABC", lines[0].Text);
        TestAssert.Equal("DEFG", lines[1].Text);
        TestAssert.True(lines[0].EndsWithIntraTokenBreak, "Zero-width spaces should be preferred hidden break opportunities inside overwide tokens.");
        TestAssert.True(
            lines[1].Segments.Any(segment => segment.SourceTextRunIndex == 8 && segment.SourceTextOffsetInRun == 4 && segment.Text == "DEFG"),
            "Hidden zero-width spaces should not shift source offsets for later rendered text in the same run.");
    }

    public static void DocxParagraphLayoutDoesNotBreakAtNoBreakHyphen()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABC\u2011DEFG", 10d, null, false, false, false, null, null)],
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
            45d,
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal("ABC\u2011DEFG", lines[0].Text);
    }

    public static void DocxParagraphLayoutWordWrapOffBreaksOverwideLatinTokens()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABCDEFGH", 10d, null, false, false, false, null, null)],
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
            WordWrap = false,
            WordWrapValue = "0"
        };
        var document = new DocxDocument(
            45d,
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("ABCDE", lines[0].Text);
        TestAssert.Equal("FGH", lines[1].Text);
        TestAssert.True(lines[0].EndsWithIntraTokenBreak, "wordWrap off should permit safe character-level line breaks for overwide Latin tokens.");
    }

    public static void DocxParagraphLayoutKeepsLineWithOverflowingTerminalSpace()
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
                      <w:pPr><w:ind w:left="3950" w:right="3950"/></w:pPr>
                      <w:r><w:t xml:space="preserve">aaaa bbbb cccc </w:t></w:r>
                    </w:p>
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

        // Word measures terminal spaces but does not wrap for them (width probe 2026-09-06: a 99.3pt token stays on one 100pt line despite its 2.03pt terminal space).
        // RV06 row-end matrix: Office keeps one row-end space beyond authored trailing.
        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal("aaaa bbbb cccc  ", lines[0].Text);
    }

    // count-based scaling tests plus R07 linear gate: deterministic unit-width measurer, line
    // width 10pt, emergency (allowOverwideTokenBreaks) wrapping. Pre-fix baselines
    // measured 2026-09-21 on this code (artifacts/wrap-baseline.txt, ignored):
    // L=128: 795 calls / 40,798 chars; L=256: 3,228 / 303,593; L=512: 13,008 /
    // 2,334,038 (matching the pre-fix probe baselines above). R07 estimates fit via average char width
    // plus local grow/shrink among safe breaks, so 128/256/512 thresholds below enforce linear
    // growth (2x per doubling, not 4x); golden lengths pin identical line breaking.
    internal sealed class CountingUnitMeasurer : IDocxTextMeasurer
    {
        public int MeasureCalls;
        public long CharsMeasured;

        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            MeasureCalls++;
            CharsMeasured += text.Length;
            return text.Length * 1d;
        }
    }

    public static void EmergencyWrapLongUnbrokenTokenBoundsWork()
    {
        AssertWrapScaling(new string((char)97, 128), 200, new[] { 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 8 });
        AssertWrapScaling(new string((char)97, 256), 400, null);
        AssertWrapScaling(new string((char)97, 512), 800, null);
    }

    public static void EmergencyWrapHyphenatedTokenBoundsWork()
    {
        var token = new System.Text.StringBuilder();
        for (int i = 0; i < 512; i++)
        {
            token.Append(i % 8 == 7 ? (char)45 : (char)98);
        }

        // 64 eight-char hyphen groups break at group ends; counts must not regress.
        AssertWrapScaling(token.ToString(), 2080, null);
    }

    public static void EmergencyWrapSoftHyphenTokenKeepsStarts()
    {
        var token = new System.Text.StringBuilder();
        for (int i = 0; i < 256; i++)
        {
            token.Append(i % 16 == 15 ? (char)0xAD : (char)99);
        }

        IReadOnlyList<DocxWrappedTextLine> lines = WrapWithCounting(token.ToString(), 10d, true, out int calls);
        TestAssert.True(calls <= 14000, $"Soft-hyphen token must stay bounded, saw {calls} measures.");
        // 16-char soft-hyphen period breaks as 10+6 lines with the hyphen preserved
        // visibly at each break; 15 visible hyphens + 240 base chars = 255 covered.
        TestAssert.Equal(32, lines.Count);
        int covered = 0;
        for (int i = 0; i < lines.Count; i++)
        {
            // The final line drops its terminal soft hyphen (no intra-token break).
            TestAssert.Equal(i == 31 ? 5 : i % 2 == 0 ? 10 : 6, lines[i].Text.Length);
            covered += lines[i].Text.Length;
        }

        TestAssert.Equal(255, covered);
    }

    public static void EmergencyWrapMultiRunTokenKeepsStarts()
    {
        string first = new string((char)97, 100);
        string second = new string((char)98, 100);
        string third = new string((char)99, 100);
        var run1 = new DocxTextRun(first, 11d, null, false, false, false, null, "A");
        var run2 = new DocxTextRun(second, 11d, null, false, false, false, null, "B");
        var run3 = new DocxTextRun(third, 11d, null, false, false, false, null, "C");
        var spans = new[]
        {
            new DocxTextSpan(first, run1, 0, 0),
            new DocxTextSpan(second, run2, 1, 0),
            new DocxTextSpan(third, run3, 2, 0),
        };
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpansWithCounting(first + second + third, spans, 10d, true, out int calls);
        TestAssert.Equal(30, lines.Count);
        foreach (DocxWrappedTextLine line in lines)
        {
            TestAssert.Equal(10, line.Text.Length);
        }

        TestAssert.True(calls <= 8600, $"Multi-run token must stay bounded, saw {calls} measures.");
    }

    public static void EmergencyWrapUnicodeBoundariesAreSafe()
    {
        // Surrogate pairs and combining marks must never split across lines.
        string token = "ab\U0001F600cde\u0301f" + new string((char)97, 60);
        IReadOnlyList<DocxWrappedTextLine> lines = WrapWithCounting(token, 10d, true, out _);
        int covered = 0;
        foreach (DocxWrappedTextLine line in lines)
        {
            TestAssert.True(line.Text.Length > 0, "No empty lines.");
            TestAssert.True(!char.IsLowSurrogate(line.Text[0]), "Line must not start inside a surrogate pair.");
            TestAssert.True(!char.IsHighSurrogate(line.Text[^1]), "Line must not end inside a surrogate pair.");
            covered += line.Text.Length;
        }

        TestAssert.Equal(token.Length, covered);
    }

    public static void EmergencyWrapShrinkCapFailsDefined()
    {
        // R07.2: ten huge stems ahead of a long tiny tail make every feasible prefix
        // overflow, so the backward scan would walk the whole token. Past 128 shrink
        // probes the conversion fails defined instead of searching unbounded.
        var token = new System.Text.StringBuilder();
        token.Append(new string((char)72, 10));
        token.Append(new string((char)97, 15000));
        var measurer = new ProfilingMeasurer((run, text) =>
        {
            double width = 0d;
            foreach (char c in text)
            {
                width += c == (char)72 ? 100d : 0.01d;
            }
            return width;
        });
        OoxPdfLimitExceededException thrown = TestAssert.Throws<OoxPdfLimitExceededException>(() => WrapSpansWithWidths(token.ToString(), SingleSpan(token.ToString()), _ => 10d, true, measurer));
        TestAssert.Contains("shrink probes", thrown.Message);
    }
    public static void EmergencyWrapCancelledTokenThrows()
    {
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        var measurer = new CountingUnitMeasurer();
        string token = new string((char)97, 512);
        var run = new DocxTextRun(token, 11d, null, false, false, false, null, "Test");
        var spans = new[] { new DocxTextSpan(token, run, 0, 0) };
        TestAssert.Throws<OperationCanceledException>(() => WrapSpans(token, spans, 10d, true, measurer, cancelled.Token));
    }

    // R07.1 characterization: work instrumentation beyond call counts. Records calls,
    // measured characters, distinct slices (a memo-effectiveness proxy), and per-run
    // attribution, so scaling tests pin work volume rather than bare call counts.
    // Shaping stays stubbed by design (deterministic unit widths); real-shaping
    // equivalence belongs to R07.2 layout semantics plus visual gates.
    internal sealed class ProfilingMeasurer(Func<DocxTextRun?, string, double> width) : IDocxTextMeasurer
    {
        public int MeasureCalls;
        public long CharsMeasured;
        public readonly HashSet<string> DistinctSlices = new(StringComparer.Ordinal);
        public readonly Dictionary<string, int> CallsByRun = new(StringComparer.Ordinal);
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            MeasureCalls++;
            CharsMeasured += text.Length;
            DistinctSlices.Add(text);
            string key = run?.FontFamily ?? "(null)";
            CallsByRun[key] = CallsByRun.TryGetValue(key, out int count) ? count + 1 : 1;
            return width(run, text);
        }
    }
    // Strongly uneven advances: narrow stems with very wide rounds defeat the
    // average-width fit estimate, exercising grow/shrink search depth.
    private static double AlternatingWidth(DocxTextRun? run, string text)
    {
        double width = 0d;
        foreach (char c in text)
        {
            width += c == (char)97 ? 0.5d : 8d;
        }
        return width;
    }
    // Negative tracking on AV pairs makes longer prefixes measure smaller, so prefix
    // widths are not monotonic and the shorter-safe-break fallback runs.
    private static double KernedWidth(DocxTextRun? run, string text)
    {
        double width = text.Length * 1d;
        int index = 0;
        while ((index = text.IndexOf("AV", index, StringComparison.Ordinal)) >= 0)
        {
            width -= 0.9d;
            index += 2;
        }
        return width;
    }
    // Run-varying advances for many-short-run scaling (run count varies at fixed length).
    private static double FamilyWidth(DocxTextRun? run, string text)
    {
        return text.Length * (run?.FontFamily == "Wide" ? 8d : 0.5d);
    }
    // R07.1: alternating 0.5/8.0 advances at 10pt capacity (383 calls / 17,022 chars
    // measured 2026-09-23). Bounds pin current work volume; fit plus coverage pin the
    // line contract under adversarial widths.
    public static void EmergencyWrapUnevenAdvancesBoundsWork()
    {
        var token = new System.Text.StringBuilder();
        for (int i = 0; i < 256; i++)
        {
            token.Append(i % 2 == 0 ? (char)97 : (char)98);
        }
        var measurer = new ProfilingMeasurer(AlternatingWidth);
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpansWithWidths(token.ToString(), SingleSpan(token.ToString()), _ => 10d, true, measurer);
        TestAssert.True(measurer.MeasureCalls <= 1200, string.Format("Uneven advances must stay bounded, saw {0} measures.", measurer.MeasureCalls));
        TestAssert.True(measurer.CharsMeasured <= 60000, string.Format("Uneven advances must bound measured characters, saw {0}.", measurer.CharsMeasured));
        AssertCoverage(token.ToString(), lines);
        AssertAllLinesFit(lines, _ => 10d, AlternatingWidth);
    }
    // R07.1: AV-kerning makes prefix widths non-monotonic at 10pt capacity (43 calls /
    // 2,468 chars measured 2026-09-23).
    public static void EmergencyWrapNegativeTrackingBoundsWork()
    {
        var token = new System.Text.StringBuilder();
        for (int i = 0; i < 128; i++)
        {
            token.Append("AV");
        }
        var measurer = new ProfilingMeasurer(KernedWidth);
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpansWithWidths(token.ToString(), SingleSpan(token.ToString()), _ => 10d, true, measurer);
        TestAssert.True(measurer.MeasureCalls <= 200, string.Format("Negative tracking must stay bounded, saw {0} measures.", measurer.MeasureCalls));
        TestAssert.True(measurer.CharsMeasured <= 10000, string.Format("Negative tracking must bound measured characters, saw {0}.", measurer.CharsMeasured));
        AssertCoverage(token.ToString(), lines);
        AssertAllLinesFit(lines, _ => 10d, KernedWidth);
    }
    // R07.1: run count varies independently at fixed 256-char length (8/16/32 runs).
    public static void EmergencyWrapManyShortRunsBoundsWork()
    {
        int calls8 = WrapFamilyRuns(8, out _);
        int calls16 = WrapFamilyRuns(16, out IReadOnlyList<DocxWrappedTextLine> lines16);
        int calls32 = WrapFamilyRuns(32, out _);
        TestAssert.True(calls16 <= 3 * calls8, string.Format("Doubling runs must scale calls linearly, saw {0} then {1}.", calls8, calls16));
        TestAssert.True(calls32 <= 3 * calls16, string.Format("Doubling runs must scale calls linearly, saw {0} then {1}.", calls16, calls32));
        foreach (DocxWrappedTextLine line in lines16)
        {
            TestAssert.True(line.Text.Length > 0, "No empty lines.");
        }
    }
    // R07.1: line capacity varies independently on one 256-char uniform token.
    public static void EmergencyWrapCapacityScalingBoundsWork()
    {
        string token = new string((char)97, 256);
        int calls10 = WrapUniformWithWidths(token, _ => 10d, out IReadOnlyList<DocxWrappedTextLine> lines10);
        int calls20 = WrapUniformWithWidths(token, _ => 20d, out IReadOnlyList<DocxWrappedTextLine> lines20);
        int calls40 = WrapUniformWithWidths(token, _ => 40d, out _);
        TestAssert.True(calls20 <= calls10, string.Format("Doubling capacity must not add work, saw {0} then {1}.", calls10, calls20));
        TestAssert.True(calls40 <= calls20, string.Format("Doubling capacity must not add work, saw {0} then {1}.", calls20, calls40));
        AssertCoverage(token, lines10);
        AssertCoverage(token, lines20);
    }
    // R07.1: N/2N/4N allocation evidence above fixed setup costs. GC bytes are exact
    // (not sampled), so ratios pin the growth shape deterministically.
    public static void EmergencyWrapAllocationScalesLinearly()
    {
        long alloc128 = WrapAllocated(new string((char)97, 128));
        long alloc256 = WrapAllocated(new string((char)97, 256));
        long alloc512 = WrapAllocated(new string((char)97, 512));
        TestAssert.True(alloc256 - alloc128 > 0, "Growth must be observable above setup.");
        TestAssert.True(alloc512 - alloc256 <= (long)(2.5d * (alloc256 - alloc128)), string.Format("Allocations must scale linearly, saw deltas {0} then {1}.", alloc256 - alloc128, alloc512 - alloc256));
    }
    // RV13: measured characters must scale linearly at thousands of characters,
    // not fourfold per doubling (per-line full-remainder measurement).
    public static void EmergencyWrapMeasuredCharactersScaleLinearly()
    {
        long chars2048 = WrapUnitChars(2048);
        long chars4096 = WrapUnitChars(4096);
        long chars8192 = WrapUnitChars(8192);
        TestAssert.True(chars4096 <= 3L * chars2048, $"Measured characters must scale linearly, saw {chars2048} then {chars4096}.");
        TestAssert.True(chars8192 <= 3L * chars4096, $"Measured characters must scale linearly, saw {chars4096} then {chars8192}.");
    }

    private static long WrapUnitChars(int length)
    {
        string token = new string((char)97, length);
        var measurer = new ProfilingMeasurer((run, text) => text.Length * 1d);
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpansWithWidths(token, SingleSpan(token), _ => 10d, true, measurer);
        TestAssert.True(lines.Count > 0, "Token must produce lines.");
        return measurer.CharsMeasured;
    }

    // RV13: allocations must scale linearly at thousands of characters.
    public static void EmergencyWrapLargeAllocationScalesLinearly()
    {
        long alloc2048 = WrapAllocated(new string((char)97, 2048));
        long alloc4096 = WrapAllocated(new string((char)97, 4096));
        long alloc8192 = WrapAllocated(new string((char)97, 8192));
        TestAssert.True(alloc2048 > 0, "Growth must be observable above setup.");
        TestAssert.True(alloc4096 <= (long)(3.5d * alloc2048), $"Allocations must scale linearly, saw {alloc2048} then {alloc4096}.");
        TestAssert.True(alloc8192 <= (long)(3.5d * alloc4096), $"Allocations must scale linearly, saw {alloc4096} then {alloc8192}.");
    }

    // RV13: wide head characters with a narrow tail defeat the chain-average
    // fit estimate, exercising the fitting-tail confirmation.
    private static double HeadWideTailNarrowWidth(DocxTextRun? run, string text)
    {
        double width = 0d;
        foreach (char c in text)
        {
            width += c == (char)72 ? 10d : c == (char)45 ? 5d : 1d;
        }
        return width;
    }

    // RV13: a mixed-width chain whose narrow tail fits must not be broken when
    // the chain-average estimate overflows. Wide head characters inflate the
    // recorded average, but the fitting emergency tail needs no break.
    public static void EmergencyWrapMixedWidthFittingTailNeedsNoBreak()
    {
        string token = new string((char)72, 10) + new string((char)108, 10);
        var measurer = new ProfilingMeasurer(HeadWideTailNarrowWidth);
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpansWithWidths(token, SingleSpan(token), _ => 10d, true, measurer);
        AssertCoverage(token, lines);
        TestAssert.Equal(11, lines.Count);
        TestAssert.Equal(new string((char)108, 10), lines[10].Text);
        AssertAllLinesFit(lines, _ => 10d, HeadWideTailNarrowWidth);
    }

    // RV13: same fitting-tail requirement through preferred (hyphen) breaks.
    public static void PreferredWrapMixedWidthFittingTailNeedsNoBreak()
    {
        string token = new string((char)72, 10) + "-" + new string((char)108, 10);
        var measurer = new ProfilingMeasurer(HeadWideTailNarrowWidth);
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpansWithWidths(token, SingleSpan(token), _ => 15d, true, measurer);
        AssertCoverage(token, lines);
        TestAssert.Equal(11, lines.Count);
        TestAssert.Equal("H-", lines[9].Text);
        TestAssert.Equal(new string((char)108, 10), lines[10].Text);
        AssertAllLinesFit(lines, _ => 15d, HeadWideTailNarrowWidth);
    }

    // R07.1: first-line (6pt) versus continuation (10pt) widths pin distinct break
    // offsets, including the ragged tail; R07.2 must preserve both widths.
    public static void WrapFirstContinuationWidthsPinStarts()
    {
        string token = new string((char)97, 128);
        var measurer = new ProfilingMeasurer((run, text) => text.Length * 1d);
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpansWithWidths(token, SingleSpan(token), lineIndex => lineIndex == 0 ? 6d : 10d, true, measurer);
        AssertCoverage(token, lines);
        TestAssert.Equal(6, lines[0].Text.Length);
        for (int i = 1; i < lines.Count - 1; i++)
        {
            TestAssert.Equal(10, lines[i].Text.Length);
        }
        TestAssert.Equal(2, lines[lines.Count - 1].Text.Length);
        TestAssert.Equal(14, lines.Count);
    }
    // R07.1: justification stretches lines but must not re-break them: a justified
    // paragraph pins the same break offsets as its left-aligned twin.
    public static void JustifiedParagraphLineStartsArePinned()
    {
        string[] justified = LayoutLineTexts(justified: true);
        string[] plain = LayoutLineTexts(justified: false);
        TestAssert.Equal(plain.Length, justified.Length);
        TestAssert.True(justified.Length > 1, "Case must wrap to several lines.");
        for (int i = 0; i < justified.Length; i++)
        {
            TestAssert.Equal(plain[i], justified[i]);
        }
    }

    private static int WrapFamilyRuns(int runs, out IReadOnlyList<DocxWrappedTextLine> lines)
    {
        int partLength = 256 / runs;
        var text = new System.Text.StringBuilder();
        var spans = new List<DocxTextSpan>();
        for (int i = 0; i < runs; i++)
        {
            string part = new string(i % 2 == 0 ? (char)97 : (char)98, partLength);
            text.Append(part);
            var run = new DocxTextRun(part, 11d, null, false, false, false, null, i % 2 == 0 ? "Narrow" : "Wide");
            spans.Add(new DocxTextSpan(part, run, i, 0));
        }
        var measurer = new ProfilingMeasurer(FamilyWidth);
        lines = WrapSpansWithWidths(text.ToString(), spans.ToArray(), _ => 10d, true, measurer);
        AssertCoverage(text.ToString(), lines);
        AssertAllLinesFit(lines, _ => 10d, FamilyWidth);
        return measurer.MeasureCalls;
    }

    private static int WrapUniformWithWidths(string token, Func<int, double> widths, out IReadOnlyList<DocxWrappedTextLine> lines)
    {
        var measurer = new ProfilingMeasurer((run, text) => text.Length * 1d);
        lines = WrapSpansWithWidths(token, SingleSpan(token), widths, true, measurer);
        AssertAllLinesFit(lines, widths, (run, text) => text.Length * 1d);
        return measurer.MeasureCalls;
    }

    private static long WrapAllocated(string token)
    {
        WrapUniformWithWidths(token, _ => 10d, out _);
        long before = GC.GetAllocatedBytesForCurrentThread();
        WrapUniformWithWidths(token, _ => 10d, out _);
        return GC.GetAllocatedBytesForCurrentThread() - before;
    }

    private static DocxTextSpan[] SingleSpan(string token)
    {
        var run = new DocxTextRun(token, 11d, null, false, false, false, null, "Test");
        return new[] { new DocxTextSpan(token, run, 0, 0) };
    }

    private static void AssertCoverage(string token, IReadOnlyList<DocxWrappedTextLine> lines)
    {
        int covered = 0;
        foreach (DocxWrappedTextLine line in lines)
        {
            TestAssert.True(line.Text.Length > 0, "No empty lines.");
            covered += line.Text.Length;
        }
        TestAssert.Equal(token.Length, covered);
    }

    private static void AssertAllLinesFit(IReadOnlyList<DocxWrappedTextLine> lines, Func<int, double> widths, Func<DocxTextRun?, string, double> width)
    {
        var verifier = new ProfilingMeasurer(width);
        for (int i = 0; i < lines.Count; i++)
        {
            double lineWidth = 0d;
            foreach (DocxTextSpan span in lines[i].Spans)
            {
                lineWidth += verifier.MeasureText(span.StyleRun, span.Text, 11d);
            }
            TestAssert.True(lineWidth <= widths(i) + 1e-9, string.Format("Line {0} must fit: {1} exceeds {2}.", i, lineWidth, widths(i)));
        }
    }
    private static void AssertWrapScaling(string token, int maxCalls, int[]? goldenLengths)
    {
        IReadOnlyList<DocxWrappedTextLine> lines = WrapWithCounting(token, 10d, true, out int calls);
        TestAssert.True(calls <= maxCalls, $"Token length {token.Length} must stay within {maxCalls} measures, saw {calls}.");
        if (goldenLengths is not null)
        {
            TestAssert.Equal(goldenLengths.Length, lines.Count);
            for (int i = 0; i < goldenLengths.Length; i++)
            {
                TestAssert.Equal(goldenLengths[i], lines[i].Text.Length);
            }
        }

        int covered = 0;
        foreach (DocxWrappedTextLine line in lines)
        {
            covered += line.Text.Length;
        }

        TestAssert.Equal(token.Length, covered);
    }

    private static IReadOnlyList<DocxWrappedTextLine> WrapWithCounting(string token, double width, bool allowOverwide, out int calls)
    {
        var run = new DocxTextRun(token, 11d, null, false, false, false, null, "Test");
        var spans = new[] { new DocxTextSpan(token, run, 0, 0) };
        return WrapSpansWithCounting(token, spans, width, allowOverwide, out calls);
    }

    private static IReadOnlyList<DocxWrappedTextLine> WrapSpansWithCounting(string token, DocxTextSpan[] spans, double width, bool allowOverwide, out int calls)
    {
        var measurer = new CountingUnitMeasurer();
        IReadOnlyList<DocxWrappedTextLine> lines = WrapSpans(token, spans, width, allowOverwide, measurer, CancellationToken.None);
        calls = measurer.MeasureCalls;
        return lines;
    }

    private static string[] LayoutLineTexts(bool justified)
    {
        string alignment = justified ? "<w:jc w:val=\"both\"/>" : string.Empty;
        string body = "<w:p><w:pPr>" + alignment + "<w:ind w:left=\"5000\" w:right=\"5000\"/></w:pPr>"
            + "<w:r><w:t xml:space=\"preserve\">alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu</w:t></w:r></w:p>";
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>",
            ["_rels/.rels"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>",
            ["word/document.xml"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:body>" + body + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" + "</w:body></w:document>",
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Select(line => line.Text)
            .ToArray();
    }
    private static IReadOnlyList<DocxWrappedTextLine> WrapSpans(string token, DocxTextSpan[] spans, double width, bool allowOverwide, IDocxTextMeasurer measurer, CancellationToken cancellationToken)
    {
        return WrapSpansWithWidths(token, spans, _ => width, allowOverwide, measurer, cancellationToken);
    }
    private static IReadOnlyList<DocxWrappedTextLine> WrapSpansWithWidths(string token, DocxTextSpan[] spans, Func<int, double> widths, bool allowOverwide, IDocxTextMeasurer measurer, CancellationToken cancellationToken = default)
    {
        MethodInfo wrap = typeof(DocxLayoutEngine).GetMethod("WrapWords", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected WrapWords.");
        try
        {
            object? result = wrap.Invoke(null, [token, spans, 0, token.Length, widths, 11d, measurer, Array.Empty<DocxTabStop>(), 36d, allowOverwide, null, cancellationToken, null]);
            return ((System.Collections.IEnumerable)result!).Cast<DocxWrappedTextLine>().ToArray();
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }
}
