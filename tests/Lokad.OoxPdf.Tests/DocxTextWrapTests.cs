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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "AlphaBeta".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

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
        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal("aaaa bbbb cccc ", lines[0].Text);
    }
}
