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

internal static class DocxTableCellsTests
{
    public static void DocxSyntheticThreeDTableBorderStylesRenderWithoutDiagnostic()
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
                              <w:top w:val="threeDEmboss" w:color="336699" w:sz="16"/>
                              <w:bottom w:val="threeDEngrave" w:color="336699" w:sz="16"/>
                            </w:tcBorders>
                          </w:tcPr>
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
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 4, "3D table borders should render light/dark relief strips instead of flattening to one solid fill.");
        TestAssert.True(pdf.Split(" rg", StringSplitOptions.None).Length - 1 >= 4, "3D table borders should emit separate light and dark fill colors.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(2, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticWaveTableBorderStylesRenderWithoutDiagnostic()
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
                              <w:top w:val="wave" w:color="336699" w:sz="16"/>
                              <w:bottom w:val="doubleWave" w:color="336699" w:sz="16"/>
                            </w:tcBorders>
                          </w:tcPr>
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
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(DocxTests.CountOccurrences(pdf, " l S") >= 8, "Wave table borders should render as repeated stroked wave segments.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(2, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticTripleTableBorderStyleRendersWithoutDiagnostic()
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
                              <w:top w:val="triple" w:color="993333" w:sz="18"/>
                            </w:tcBorders>
                          </w:tcPr>
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
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.6 0.2 0.2 rg", pdf);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 3, "Triple borders should render as three filled strips.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(1, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticCompoundTableBorderStylesRenderWithoutDiagnostic()
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
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="thinThickSmallGap" w:color="225588" w:sz="24"/>
                              <w:bottom w:val="thickThinSmallGap" w:color="225588" w:sz="24"/>
                              <w:left w:val="thinThickThinSmallGap" w:color="225588" w:sz="24"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="thinThickMediumGap" w:color="225588" w:sz="24"/>
                              <w:bottom w:val="thickThinMediumGap" w:color="225588" w:sz="24"/>
                              <w:left w:val="thinThickThinMediumGap" w:color="225588" w:sz="24"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="thinThickLargeGap" w:color="225588" w:sz="24"/>
                              <w:bottom w:val="thickThinLargeGap" w:color="225588" w:sz="24"/>
                              <w:right w:val="thinThickThinLargeGap" w:color="225588" w:sz="24"/>
                            </w:tcBorders>
                          </w:tcPr>
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
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 18, "Compound table borders should render as separate thin/thick filled strips.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(9, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.Equal("top", cells[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("center", cells[1].VerticalAlignmentValue ?? string.Empty);
        TestAssert.True(cells[2].VerticalAlignmentValue is null, "Expected missing cell vertical alignment to keep a null source token.");
    }

    public static void DocxReaderTableCellPreservesTextDirectionToken()
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
                        <w:tc><w:tcPr><w:textDirection w:val="tbRl"/></w:tcPr><w:p><w:r><w:t>Rotated</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Default</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.Equal("tbRl", cells[0].TextDirectionValue ?? string.Empty);
        TestAssert.True(cells[1].TextDirectionValue is null, "Expected missing cell text direction to keep a null source token.");
    }

    public static void DocxReaderTableCellPreservesNoWrapToken()
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
                        <w:tc><w:tcPr><w:noWrap/></w:tcPr><w:p><w:r><w:t>Kept</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:noWrap w:val="0"/></w:tcPr><w:p><w:r><w:t>Wrapped</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.True(cells[0].NoWrap, "Expected implicit w:noWrap to disable automatic cell text wrapping.");
        TestAssert.True(cells[0].NoWrapValue is null, "Expected implicit w:noWrap to preserve a null source token.");
        TestAssert.True(!cells[1].NoWrap, "Expected w:noWrap w:val=\"0\" to allow automatic cell text wrapping.");
        TestAssert.Equal("0", cells[1].NoWrapValue ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesFitTextToken()
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
                        <w:tc><w:tcPr><w:tcFitText/></w:tcPr><w:p><w:r><w:t>Fit</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:tcFitText w:val="0"/></w:tcPr><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.True(cells[0].FitText, "Expected implicit w:tcFitText to fit cell text to its extents.");
        TestAssert.True(cells[0].FitTextValue is null, "Expected implicit w:tcFitText to preserve a null source token.");
        TestAssert.True(!cells[1].FitText, "Expected w:tcFitText w:val=\"0\" to disable cell text fitting.");
        TestAssert.Equal("0", cells[1].FitTextValue ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableCellMargins margins = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0].Cells[0].Margins;

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTable table = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0];

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        TestAssert.True(DocxTests.CountOccurrences(pdf, "/PatternType 1") >= shadingValues.Length, "Each supported DOCX stripe family should have a distinct tiling pattern resource.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, diagnostics.Add, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];

        TestAssert.Equal("1.", cell.Paragraphs[0].ListLabel?.Text ?? string.Empty);
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "1. Item".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);
        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
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
        int firstText = DocxTests.FirstPdfTextShowIndex(pdf);
        int tableGrid = pdf.IndexOf(" re f", StringComparison.Ordinal);
        int lastText = DocxTests.LastPdfTextShowIndex(pdf);
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
        // Table terminus (w7 doc-start probe): the last-row bottom border hangs below
        // content, so the single-row bottom strip sits one 0.48pt width lower than the
        // legacy bottom-inside convention placed it.
        // w63 centered-vertical law: the strip starts at grid + half border width (72.24); the end keeps the butt-overlap convention.
        TestAssert.Contains("72.24 683.04 143.76 0.48 re f", pdf);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableRow row = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0];
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableRow row = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0];

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableRow row = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0];

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
        DocxTable first = DocxTests.CreateSingleCellTable("first", rowHeight: 20d);
        DocxTable second = DocxTests.CreateSingleCellTable("second", rowHeight: 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxPageBreakElement(DocxBreakSourceKind.PageBreakBefore, null, null),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageManualColumnBreakAdvancesPageInSingleColumn()
    {
        DocxTable first = DocxTests.CreateSingleCellTable("first", rowHeight: 20d);
        DocxTable second = DocxTests.CreateSingleCellTable("second", rowHeight: 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxManualBreakElement(DocxBreakSourceKind.RunBreak, "column", null),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

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
                new DocxSectionBreakElement(sectionSettings, DocxSectionBreakType.NextPage, "2", "1", "360", [])
            ],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
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
        DocxTable first = DocxTests.CreateSingleCellTable("first", rowHeight: 150d);
        DocxTable second = DocxTests.CreateSingleCellTable("second", rowHeight: 20d);
        DocxParagraph marker = DocxTests.CreateDocxLayoutParagraph("marker", 10d, 20d);
        DocxParagraph breakParagraph = DocxTests.CreateDocxLayoutParagraph("", 10d, 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxParagraphElement(marker),
            new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", breakParagraph),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(0, layout.Pages[1].Items.Count);
        TestAssert.Equal(1, layout.Pages[2].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageRepeatsHeaderRowsAfterPageBreak()
    {
        var header = new DocxTableRow([new DocxTableCell("Header", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d) with {IsHeader = true };
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        TestAssert.Equal(2, firstPageRows.Length);
        TestAssert.Equal(2, secondPageRows.Length);
        TestAssert.Equal("Header", secondPageRows[0].Cells[0].Cell.Text);
        TestAssert.Equal("Second", secondPageRows[1].Cells[0].Cell.Text);
    }

    public static void DocxTableRowHeightIncludesListLabelFirstLineExtraLeading()
    {
        DocxListLabel tallLabel = new DocxListLabel("*", "bullet", "*", "tab", "1", 0, DocxNumberingIndent.Empty, new DocxTextRunStyle(10d, null, false, false, false, null, "Label Metrics", new DocxRunFonts("Label Metrics", null, null, null, null, null, null, null)));
        DocxListLabel flatLabel = new DocxListLabel("*", "bullet", "*", "tab", "1", 0, DocxNumberingIndent.Empty, new DocxTextRunStyle(10d, null, false, false, false, null, null, new DocxRunFonts(null, null, null, null, null, null, null, null)));
        DocxParagraphSpacing autoSpacing = new DocxParagraphSpacing(null, null, null, null, null, null, null, "auto", null);
        DocxParagraph tallItem = new DocxParagraph([new DocxTextRun("Tall", 10d, null, false, false, false, null, null)], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1.15d, null, autoSpacing, DocxParagraphKeepRules.Empty, tallLabel);
        DocxParagraph flatItem = new DocxParagraph([new DocxTextRun("Flat", 10d, null, false, false, false, null, null)], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1.15d, null, autoSpacing, DocxParagraphKeepRules.Empty, flatLabel);
        var tallRow = new DocxTableRow([new DocxTableCell("Tall", [tallItem], null, null, null, null, [], DocxTableCellMargins.Empty)], null);
        var flatRow = new DocxTableRow([new DocxTableCell("Flat", [flatItem], null, null, null, null, [], DocxTableCellMargins.Empty)], null);
        var table = new DocxTable(null, [400d], [tallRow, flatRow]);
        var document = new DocxDocument(400d, 400d, 10d, 10d, 10d, 10d, DocxPageSettings.Empty, [], [], [], [new DocxTableElement(table)], [], [table]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] tallRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        TestAssert.Equal(2, tallRows.Length);
        TestAssert.Equal(12.6d, Math.Round(tallRows[0].Height, 4));
        TestAssert.Equal(11.6d, Math.Round(tallRows[1].Height, 4));
    }


    public static void DocxTableCellPageFieldsEvaluatePerFragmentPage()
    {
        // W04: PAGE fields must evaluate against each fragment page. The static
        // cell memo omits page args only for cells without dynamic field runs, so
        // the two fragments below must resolve different page numbers.
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] secondParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph(
                index == 4 ? "Page {PAGE} mark" : "Line " + index.ToString(CultureInfo.InvariantCulture),
                10d,
                10d))
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        string firstFragmentText = string.Concat(layout.Pages[0].Items.OfType<DocxTableRowLayout>().Where(row => row.RowIndex == 1).SelectMany(row => row.Cells[0].TextLines).Select(line => line.Text));
        string secondFragmentText = string.Concat(layout.Pages[1].Items.OfType<DocxTableRowLayout>().Where(row => row.RowIndex == 1).SelectMany(row => row.Cells[0].TextLines).Select(line => line.Text));
        bool firstHasOwn = firstFragmentText.Contains("Page 1 mark", StringComparison.Ordinal);
        bool secondHasOwn = secondFragmentText.Contains("Page 2 mark", StringComparison.Ordinal);
        bool firstHasForeign = firstFragmentText.Contains("Page 2 mark", StringComparison.Ordinal);
        bool secondHasForeign = secondFragmentText.Contains("Page 1 mark", StringComparison.Ordinal);
        TestAssert.True(
            (firstHasOwn && !firstHasForeign && !secondHasOwn && !secondHasForeign) ||
            (!firstHasOwn && !firstHasForeign && secondHasOwn && !secondHasForeign),
            "Exactly one split fragment must resolve PAGE against its own page. First: <" + firstFragmentText + "> Second: <" + secondFragmentText + ">");
        TestAssert.True(!firstFragmentText.Contains("{PAGE}", StringComparison.Ordinal) && !secondFragmentText.Contains("{PAGE}", StringComparison.Ordinal), "Split fragments must not leak unresolved PAGE placeholders.");
    }

    public static void DocxTableCellTextLinesMemoSharesSplitCells()
    {
        // W04: the feasibility and final layouts of a split static cell must hit
        // the shared memo instead of recomputing. A zero hit count means the key
        // partitions every lookup and the memo is dead weight.
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] secondParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
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

        DocxLayoutEngine.DocxTableCellTextLinesMemo.ResetTotals();
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.True(DocxLayoutEngine.DocxTableCellTextLinesMemo.TotalHits > 0, "Split static cells must share memoized text lines.");
    }

    public static void DocxTableLayoutStageSplitsTallRowsAcrossPagesByDefault()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] secondParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
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

    public static void DocxTableLayoutStageKeepsCellKeepLinesParagraphWholeAtPageBoundary()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph keptCellParagraph = DocxTests.CreateDocxLayoutParagraph(
            string.Join('\n', Enumerable.Range(1, 8).Select(index => "Line " + index.ToString(CultureInfo.InvariantCulture))),
            10d,
            10d,
            keepRules: DocxParagraphKeepRules.Empty with { KeepLines = true });
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [keptCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 80d);
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(0, secondPageRows[0].FragmentIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentCount);
        TestAssert.Equal("None", secondPageRows[0].FragmentReason);
        TestAssert.Equal(8, secondPageRows[0].Cells[0].TextLines.Count);
        TestAssert.Equal(0, DocxLayoutSnapshot.FromLayout(layout).Tables.Single().FragmentedRowCount);
    }

    public static void DocxTableLayoutStageKeepsCellWidowControlledParagraphWholeAtPageBoundary()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph widowControlledParagraph = DocxTests.CreateDocxLayoutParagraph("One\nTwo\nThree", 10d, 10d);
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [widowControlledParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d);
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentCount);
        TestAssert.Equal("None", secondPageRows[0].FragmentReason);
        TestAssert.Equal(3, secondPageRows[0].Cells[0].TextLines.Count);
    }

    public static void DocxTableLayoutStageKeepsCellKeepNextParagraphWithFollowingParagraphAtPageBoundary()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph keepNextParagraph = DocxTests.CreateDocxLayoutParagraph(
            "One\nTwo",
            10d,
            10d,
            keepRules: DocxParagraphKeepRules.Empty with { KeepNext = true });
        DocxParagraph followingParagraph = DocxTests.CreateDocxLayoutParagraph("Three", 10d, 10d);
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [keepNextParagraph, followingParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d);
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentCount);
        TestAssert.Equal("None", secondPageRows[0].FragmentReason);
        TestAssert.Equal(3, secondPageRows[0].Cells[0].TextLines.Count);
    }

    public static void DocxSharedHorizontalTableBorderHangsBelowBoundary()
    {
        // Office A/B (w9 exact/atLeast bordered row pairs, Word-COM rendered plus
        // PdfInspect, re-rendered digit-identical): interior horizontal border bands
        // hang below the row boundary by the full nominal width (atLeast mid band at
        // 683.02 for boundary 683.52, exact mid band at 683.50 for boundary 684.0),
        // while the layout centered them on the boundary (atLeast 683.28, exact
        // 683.76). Text baselines are identical both sides, so the shift is
        // pitch-neutral paint. The band start (72.24 here versus Office 72.26) is now Office-exact per the w63 centered-vertical law; the end keeps the butt-overlap convention (239.76 here versus Office 239.54).
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
                        <w:tblW w:w="0" w:type="auto"/>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideH w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideV w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid>
                      <w:tr><w:trPr><w:trHeight w:val="720" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>Row one</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:trPr><w:trHeight w:val="720" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>Row two</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>After pair</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72.24 683.02 239.76 0.48 re f", pdf);
    }

    private static string WriteW68TablePackage(string tblPrExtra)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders>
                """ + tblPrExtra + """
                </w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>S4 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>S4 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """
        });
    }

    public static void DocxTableCellTextStartsAtMaxBorderMarginInset()
    {
        // Office A/B (w63/w65 border-size probes, Word-COM rendered plus PdfInspect):
        // first-column text starts at grid + max(borderHalf, margin) with unset
        // margins defaulting to 0.48pt: sz4 text at 72.504 for grid 72.024 while
        // the layout stacked margin + borderHalf, placing sz4 text at 72.24 here.
        string input = WriteW68TablePackage("");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 72.48 709.18 Tm", pdf);
    }

    public static void DocxTableVerticalBordersCenterOnGridLines()
    {
        // Office A/B (w63 border-size probe, Word-COM rendered plus PdfInspect):
        // vertical border bands center on grid lines: sz4 outer-left at 71.784
        // for grid 72.024 with top bands butting from grid + half width, while
        // the renderer hung every band right of its edge (72.0 here).
        string input = WriteW68TablePackage("");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("71.76 704.387 0.48 15.613 re f", pdf);
        TestAssert.DoesNotContain("72 704.387 0.48 15.613 re f", pdf);
        TestAssert.Contains("72.24 719.52 119.76 0.48 re f", pdf);
    }

    public static void DocxIndentedTablePinsTextAtIndentOrigin()
    {
        // Office A/B (w62/w66/w68 indent probes, Word-COM rendered plus PdfInspect):
        // tblInd pins first-column TEXT at margin + indent for every border size
        // (w66 B4/D12/C24 all at 108.02), so the grid hangs left by
        // max(borderHalf, margin). The layout added indent without shifting,
        // placing sz4 indented text at 108.24 here.
        string input = WriteW68TablePackage("<w:tblInd w:w=\"720\" w:type=\"dxa\"/>");

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 108 709.18 Tm", pdf);
    }

    public static void DocxCompatGridAlignsOuterBorderAtMargin()
    {
        // Office A/B (m13 bare+compat15 probe, Word-COM rendered plus PdfInspect):
        // a present compatibilityMode aligns the outer border edge at the margin,
        // so the sz4 grid sits at 72.24 and text at 72.72 (modern grid would give
        // band 71.76 and text 72.48).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>C15 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>C15 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>

                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 72.72 709.18 Tm", pdf);
        TestAssert.Contains("72 704.387 0.48 15.613 re f", pdf);
    }

    public static void DocxCompatIndentAlignsOuterBorderAtIndentedMargin()
    {
        // Office A/B (m14 compat+indent probe, Word-COM rendered plus PdfInspect):
        // compatibilityMode aligns the outer border edge at margin+indent with text
        // at grid+max (Office col1 text 108.74, band [108.02, 108.50], top strip
        // butting at 108.50; pipeline unsnapped: grid 108.24, text 108.72).
        // Pre-change the unshifted grid stacked margin+border, giving 108.24 here.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders><w:tblInd w:w="720" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>M14 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>M14 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>

                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 108.72 709.18 Tm", pdf);
        TestAssert.Contains("108 704.387 0.48 15.613 re f", pdf);
    }
    public static void DocxStyleMarginsPinTextAtMarginWithoutCompat()
    {
        // Office A/B (w72 bare+styles probe, Word-COM rendered plus PdfInspect):
        // without compat, table-style margins pin text at the margin (grid hangs
        // left by the style margin: TN-108dxa grid at 66.6, text at 72.0).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>P72 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>P72 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 72 709.18 Tm", pdf);
        TestAssert.Contains("66.36 704.387 0.48 15.613 re f", pdf);
    }

    public static void DocxTableNormalMissingCellMarginsFallBackToBuiltIn()
    {
        // Office A/B (m1 ladder mutant, Word-COM rendered plus PdfInspect): a present
        // TableNormal style without stored cell margins still contributes the built-in
        // 108dxa margins (Word col1 text at 77.664 for grid 72.264 under compat;
        // pipeline emits 77.64 = 72 + 0.24 + 5.4, i.e. the same law on our unsnapped
        // base: Word export shifts every X by a global +0.024 snap, inside gate tolerance).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>F6 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>F6 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/></w:tblPr></w:style></w:styles>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 77.64 709.18 Tm", pdf);
    }

    public static void DocxDirectMarginsReplaceStyleMarginsForTextOffset()
    {
        // Office A/B (w74 ladder+direct-20dxa probe, Word-COM rendered plus PdfInspect):
        // direct margins cascade-replace style margins for the text offset (col1 at
        // 73.224 for grid 72.264, i.e. grid + max(borderHalf, direct);
        // pipeline emits 73.24 = 72 + 0.24 + max(0.24, 1.0) on our unsnapped base
        // (same +0.024 Word export snap; residual direct-margin 600dpi px-quantization
        // mapped over 8 sweep points stays parked per precedent).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders><w:tblCellMar><w:left w:w="20" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>G74 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>G74 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 73.24 709.18 Tm", pdf);
    }

    public static void DocxTableRendererDoesNotDrawRowEdgeBordersAtSplitFragmentBoundaries()
    {
        DocxTableCellBorder[] borders =
        [
            new("top", "single", "000000", "4"),
            new("bottom", "single", "000000", "4"),
            new("left", "single", "000000", "4"),
            new("right", "single", "000000", "4")
        ];
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var split = new DocxTableRow([new DocxTableCell("Split", splitParagraphs, null, null, null, null, borders, DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [first, split]);
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

        IReadOnlyList<PdfPage> pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None);

        TestAssert.Equal(2, pages.Count);
        TestAssert.DoesNotContain("10.48 10 59.52 0.48 re f", pages[0].Content);
        TestAssert.DoesNotContain("10.48 89.52 59.52 0.48 re f", pages[1].Content);
        // w63 centered-vertical law: the outer-left band centers on the grid line (9.76 for grid 10).
        TestAssert.Contains("9.76 10 0.48 20 re f", pages[0].Content);
        // Table terminus (w7 doc-start probe): the split row is the last row, so its
        // continuation fragment carries the extra bottom width and its bottom strip
        // sits 0.48pt lower; fragment-edge suppression at the split itself is unchanged.
        // w63 centered-vertical law: the bottom strip starts at grid + half border width (10.24).
        TestAssert.Contains("10.24 29.04 59.76 0.48 re f", pages[1].Content);
    }

    public static void DocxTableLayoutStageRepeatsHeaderRowsBeforeSplitRowContinuations()
    {
        DocxParagraph headerParagraph = DocxTests.CreateDocxLayoutParagraph("Header", 10d, 10d);
        DocxParagraph fillerParagraph = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var header = new DocxTableRow([new DocxTableCell("Header", [headerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 10d) with {IsHeader = true };
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
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
        DocxParagraph fillerParagraph = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 7)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
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
        DocxParagraph fillerParagraph = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
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
                DocxTableCellMargins.Empty) with {HasVerticalMerge = true,VerticalMergeValue = "restart" }
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
                DocxTableCellMargins.Empty) with {HasVerticalMerge = true }
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

        DocxTableRowLayout[] splitFragments = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cell = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        var second = new DocxTableRow([new DocxTableCell("Second", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 80d) with {CantSplit = true,CantSplitValue = "1" };
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table), new DocxParagraphElement(paragraph)], [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxTableRowLayout row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        DocxTextLineLayout following = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single();
        double expectedBaselineY = row.Y - DocxLineMetrics.ResolveBodyBaselineOffset(10d, 10d, hasExplicitLineSpacing: true);
        TestAssert.Equal(Math.Round(expectedBaselineY, 3), Math.Round(following.BaselineY, 3));
    }
}
