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

internal static class DocxTablesTests
{
    public static void DocxReaderAddsImplicitParagraphOnlyAfterTerminalBodyTable()
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
                      <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>
                          <w:tbl>
                            <w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>
                            <w:tr>
                              <w:tc>
                                <w:tcPr><w:tcW w:w="1200" w:type="dxa"/></w:tcPr>
                                <w:p><w:r><w:t>Nested</w:t></w:r></w:p>
                              </w:tc>
                            </w:tr>
                          </w:tbl>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(2, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxTableElement, "The authored terminal table should remain a table body element.");
        TestAssert.True(document.BodyElements[1] is DocxImplicitParagraphElement, "A terminal body table should synthesize Word's implicit final paragraph mark.");
        var implicitParagraph = (DocxImplicitParagraphElement)document.BodyElements[1];
        TestAssert.Equal(DocxBreakSourceKind.TerminalTable, implicitParagraph.SourceKind);

        DocxTable bodyTable = ((DocxTableElement)document.BodyElements[0]).Table;
        DocxTableCell outerCell = bodyTable.Rows.Single().Cells.Single();
        TestAssert.Equal(1, outerCell.BodyElements.Count);
        TestAssert.True(outerCell.BodyElements[0] is DocxTableElement, "Nested table-cell flow should preserve its authored terminal table.");
        TestAssert.True(!outerCell.BodyElements.OfType<DocxImplicitParagraphElement>().Any(), "The implicit terminal body paragraph is a document-body rule, not a table-cell body rule.");
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
                      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Yu Gothic"/><a:cs typeface="Arial"/></a:majorFont>
                      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Meiryo"/><a:cs typeface="Arial"/></a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxFontTableEntry entry = document.FontCatalog.Entries.Single();
        TestAssert.Equal("Corporate Sans", entry.Name);
        TestAssert.Equal("Aptos", entry.AlternateName ?? string.Empty);
        TestAssert.Equal("swiss", entry.FamilyValue ?? string.Empty);
        TestAssert.Equal("variable", entry.PitchValue ?? string.Empty);
        TestAssert.Equal("020B0604020202020204", entry.PanoseValue ?? string.Empty);
        TestAssert.Equal("00", entry.CharsetValue ?? string.Empty);
        TestAssert.Equal("Aptos Display", document.FontCatalog.ThemeFonts.MajorLatinTypeface ?? string.Empty);
        TestAssert.Equal("Aptos", document.FontCatalog.ThemeFonts.MinorLatinTypeface ?? string.Empty);
        TestAssert.Equal("Yu Gothic", document.FontCatalog.ThemeFonts.MajorEastAsiaTypeface ?? string.Empty);
        TestAssert.Equal("Meiryo", document.FontCatalog.ThemeFonts.MinorEastAsiaTypeface ?? string.Empty);
        TestAssert.Equal("Arial", document.FontCatalog.ThemeFonts.MajorComplexScriptTypeface ?? string.Empty);
        TestAssert.Equal("Arial", document.FontCatalog.ThemeFonts.MinorComplexScriptTypeface ?? string.Empty);
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
        DocxDocument document = DocxTests.CreateFontPlanDocument(
            run,
            new DocxFontCatalog(
                [new DocxFontTableEntry("Corporate Sans", "Installed Sans", "swiss", null, null, null)],
                new DocxThemeFonts("Theme Display", "Theme Sans", null, null, null, null)));
        var resolver = new MapFontResolver(["Installed Sans", "Theme Sans"], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Single();

        TestAssert.Equal(DocxTypefaceResolutionSource.FontTableAlternate, resolved.Source);
        TestAssert.Equal("Installed Sans", resolved.RequestedFamily ?? string.Empty);
        TestAssert.Equal("Installed Sans", resolved.ResolvedFamily ?? string.Empty);
        TestAssert.Equal("Corporate Sans|Installed Sans|Theme Sans", string.Join("|", resolved.CandidateFamilies));
    }

    public static void DocxFontPlanIncludesPlainTableCellText()
    {
        DocxTable table = DocxTests.CreateSingleCellTable("Cell text", 12d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        var resolver = new MapFontResolver(["Body Sans"], "Resolver Fallback");

        DocxResolvedRunTypeface resolved = DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Single();

        TestAssert.Equal("Cell text", resolved.Run.Text);
        TestAssert.Equal(DocxTypefaceResolutionSource.ResolverFallback, resolved.Source);
        TestAssert.Equal(0, resolved.CandidateFamilies.Count);
        TestAssert.Equal(DocxRenderer.DefaultDocumentTypefaceRequest, resolved.RequestedFamily ?? string.Empty);
    }

    public static void DocxBlockTraversalAndFontPlanIncludeNestedTableCellBody()
    {
        DocxParagraph nestedParagraph = DocxTests.CreateDocxLayoutParagraph("Nested cell", 12d, 12d);
        var nestedTable = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxTableCell outerCell = new(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var outerTable = new DocxTable(null, [60d], [new DocxTableRow([outerCell], 30d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(outerTable)], [outerTable]);
        var resolver = new MapFontResolver(["Arial"], "Resolver Fallback");

        DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(document).ToArray();
        DocxTable[] tables = DocxBlockTraversal.EnumerateBodyTables(document).ToArray();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);

        TestAssert.Equal(1, paragraphs.Length);
        TestAssert.Equal("Nested cell", paragraphs[0].Runs.Single().Text);
        TestAssert.Equal(2, tables.Length);
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Nested cell"), "Nested table-cell body text should participate in DOCX font planning.");
    }

    public static void DocxStructureSnapshotReportsPreLayoutBlockAndTableFacts()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("A1 body.", fontSize: 10d, lineSpacingPoints: 12d) with
        {
            StyleId = "BodyStyle",
            Indent = new DocxParagraphIndent(12d, 3d, null, 6d, "240", "60", null, "120"),
            TabStops = [new DocxTabStop(36d, "720", "left", null)],
            SnapToGrid = true,
            SnapToGridValue = "1"
        };
        DocxParagraph breakParagraph = DocxTests.CreateDocxLayoutParagraph("Break", fontSize: 9d, lineSpacingPoints: 11d);
        var listLabel = new DocxListLabel(
            "1",
            "decimal",
            "%1.",
            "tab",
            "7",
            0,
            new DocxNumberingIndent(36d, null, 6d, 18d, 54d, "720", null, "120", "360", "num", "1080"),
            DocxTextRunStyle.Empty);
        DocxParagraph cellParagraph = DocxTests.CreateDocxLayoutParagraph("Cell 42", fontSize: 9d, lineSpacingPoints: 10d) with
        {
            StyleId = "CellStyle",
            ListLabel = listLabel,
            Indent = new DocxParagraphIndent(48d, null, null, 24d, "960", null, null, "480"),
            TabStops = [new DocxTabStop(72d, "1440", "num", null)]
        };
        var restartCell = new DocxTableCell(
            string.Empty,
            [],
            "D9EAF7",
            "clear",
            null,
            "center",
            [new DocxTableCellBorder("bottom", "single", "auto", "4")],
            DocxTableCellMargins.Empty)
        {
            GridSpan = 2, GridSpanValue = "2", HasVerticalMerge = true, VerticalMergeValue = "restart",
            BodyElements = [new DocxParagraphElement(cellParagraph)]
        };
        var continuationCell = new DocxTableCell(
            string.Empty,
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true,VerticalMergeValue = "continue" };
        var table = new DocxTable(
            "fixed",
            [20d, 20d],
            [
                new DocxTableRow([restartCell], 18d) with {IsHeader = true,HeaderValue = "1",HeightValue = "360",HeightRuleValue = "atLeast",CantSplit = true,CantSplitValue = "1" },
                new DocxTableRow([continuationCell], null)
            ]) with {StyleId = "TableGrid",PreferredWidthPoints = 40d,PreferredWidthValue = "800",PreferredWidthType = "dxa",IndentPoints = 6d,IndentValue = "120",IndentType = "dxa",CellSpacingPoints = 1d,CellSpacingValue = "20",CellSpacingType = "dxa",Look = new DocxTableLook(null, true, "1", null, null, true, "1", null, null, null, null, null, null) };
        DocxParagraph documentHeader = DocxTests.CreateDocxLayoutParagraph("Header", fontSize: 9d, lineSpacingPoints: 10d);
        DocxParagraph documentFooter = DocxTests.CreateDocxLayoutParagraph("Footer", fontSize: 9d, lineSpacingPoints: 10d);
        DocxParagraph sectionHeader = DocxTests.CreateDocxLayoutParagraph("Section", fontSize: 9d, lineSpacingPoints: 10d);
        DocxPageSettings sectionSettings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["first"] = [sectionHeader]
            }
        };
        var sectionBreak = new DocxSectionBreakElement(sectionSettings, DocxSectionBreakType.NextPage, "2", "1", "720", []);
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
            DocxFloatingWrapKind.Square,
            "bothSides", ImageRelationshipId: null, Image: null, SourceParagraphIndex: null, SourceBlockIndex: null);
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
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", breakParagraph),
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

        DocxStructureSnapshot snapshot = new DocxRenderer(new MapFontResolver([], "Fallback"), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);

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
        TestAssert.Equal(1, listUsage.LeftIndentParagraphCount);
        TestAssert.Equal(0, listUsage.RightIndentParagraphCount);
        TestAssert.Equal(1, listUsage.FirstLineIndentParagraphCount);
        TestAssert.Equal(1, listUsage.HangingIndentParagraphCount);
        TestAssert.Equal(1, listUsage.NumberingTabParagraphCount);
        TestAssert.Equal(1, listUsage.ParagraphIndentOverrideCount);
        TestAssert.Equal(1, listUsage.ParagraphNumberingTabStopCount);
        DocxStructureFloatingDrawingSnapshot drawingSnapshot = snapshot.FloatingDrawings.Single();
        TestAssert.Equal("wrapSquare", drawingSnapshot.WrapKind ?? string.Empty);
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
        DocxTable table = DocxTests.CreateSingleCellTable("Plain 123", 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);

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
        DocxParagraph cellParagraph = DocxTests.CreateDocxLayoutParagraph("External Internal", fontSize: 10d, lineSpacingPoints: 12d) with
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxStructureTableSnapshot tableSnapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/cell", annotation.Uri);
        TestAssert.True(annotation.X > document.MarginLeftPoints, "The annotation should be anchored to the placed table-cell hyperlink run.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover table-cell hyperlink text.");
    }

    public static void DocxRendererEmitsTableCellHyperlinkAnnotationsNextToFields()
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:hyperlink r:id="rIdTable"><w:r><w:t>table link</w:t></w:r></w:hyperlink></w:p></w:tc>
                        <w:tc><w:p><w:fldSimple w:instr=" REF FieldTarget "><w:r><w:t>table field</w:t></w:r></w:fldSimple></w:p></w:tc>

                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/table" TargetMode="External"/>
                </Relationships>
                """
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();
        System.Collections.Generic.List<PdfLinkAnnotation> all = page.Annotations.ToList();
        TestAssert.True(all.Any(annotation => annotation.Uri == "https://example.invalid/table"), "annots=" + System.String.Join(";", all.Select(annotation => (annotation.Uri ?? "null") + "@" + annotation.X)));
        PdfLinkAnnotation annotation = all.Single(annotation => annotation.Uri == "https://example.invalid/table");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover table-cell hyperlink text next to fields.");
    }

    public static void DocxWordCompatibleAllMarkupUsesEmittedAdvanceForTableCellHyperlinkAnnotations()
    {
        const string linkText = "LinkedWords";
        var runs = new[]
        {
            new DocxTextRun("Cell ", 10d, null, false, false, false, null, null),
            new DocxTextRun(linkText, 10d, null, false, false, false, null, null)
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
                new DocxHyperlinkSpan("rIdCell", null, null, null, "https://example.invalid/cell-word-compatible", "External", null, 1, 1, 1, 1, linkText.Length)
            ]
        };
        var table = new DocxTable(
            null,
            [160d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 24d)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxTextEmissionSegmentSnapshot linkSegment = renderer
            .InspectTextEmission(document)
            .Lines
            .SelectMany(line => line.Segments)
            .Single(segment => !segment.IsTerminalLineSpace && segment.TextLength == linkText.Length);
        PdfLinkAnnotation annotation = renderer.RenderBlankPages(document, null, CancellationToken.None).Single().Annotations.Single();

        TestAssert.True(
            Math.Abs(annotation.X - (linkSegment.X - (2.3d * 10d / 11d))) < 0.001d,
            "Word-compatible all-markup table-cell hyperlink annotations should use emitted table-cell segment x coordinates.");
        TestAssert.True(
            Math.Abs(annotation.Width - (linkSegment.AdvanceProfile.PlannedEmittedAdvance + 2 * (2.3d * 10d / 11d))) < 0.001d,
            "Word-compatible all-markup table-cell hyperlink annotations should cover the emitted glyph advance after positioned spacing.");
        TestAssert.True(
            Math.Abs(annotation.Width - linkSegment.Width) > 0.05d,
            "The regression should exercise a table-cell link whose emitted advance differs from the layout segment width.");
    }

    public static void DocxRendererKeepsSplitTableCellHyperlinkAnnotationsOnVisibleFragment()
    {
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d);
        var afterRuns = new[]
        {
            new DocxTextRun("After ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var after = new DocxParagraph(
            afterRuns,
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdSplitCell", null, null, null, "https://example.invalid/split-cell", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        PdfPage[] pages = renderer.RenderBlankPages(document, null, CancellationToken.None).ToArray();
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, pages.Length);
        TestAssert.Equal(0, pages[0].Annotations.Count);
        PdfLinkAnnotation annotation = pages[1].Annotations.Single();
        TestAssert.Equal("https://example.invalid/split-cell", annotation.Uri);
        TestAssert.True(annotation.Width > 0d && annotation.Height > 0d, "The visible split fragment should keep a non-empty hyperlink annotation.");
        TestAssert.Equal("Before", layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After Link", layout.Pages[1].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines.Single().Text);
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
        DocxParagraph targetParagraph = DocxTests.CreateDocxLayoutParagraph("Bookmark target", 10d, 12d) with
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal table-cell links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal table-cell links should resolve to a PDF page destination.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed table-cell bookmark coordinates.");
        TestAssert.True(annotation.Destination?.Top > 0d, "The destination should point to a concrete table-cell bookmark line top.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should cover the rendered table-cell internal-link text.");
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

    public static void DocxTableLayoutUsesResolvedPageFieldForCellLineBreaking()
    {
        // 26pt fixture sinks the 0.48pt Office default cell insets (w68) so the resolved field still fits one line; wrap-vs-Word symmetry gets its own probe.
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("A {PAGE} B", 10d, 10d);
        var cell = new DocxTableCell("A {PAGE} B", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [26d], [new DocxTableRow([cell], null)]);
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

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        DocxTextLineLayout line = row.Cells.Single().TextLines.Single();
        TestAssert.Equal("A 1 B", line.Text);
    }

    public static void DocxTableLayoutUsesCompactNumPagesFieldForCellLineBreaking()
    {
        // 26pt fixture sinks the 0.48pt Office default cell insets (w68) so the literal field still fits one line; wrap-vs-Word symmetry gets its own probe.
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("A {NUMPAGES} B", 10d, 10d);
        var cell = new DocxTableCell("A {NUMPAGES} B", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [26d], [new DocxTableRow([cell], null)]);
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

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        DocxTextLineLayout line = row.Cells.Single().TextLines.Single();
        TestAssert.Equal("A {NUMPAGES} B", line.Text);
    }

    public static void DocxReaderUsesNoteReferenceSettingsForBodyAndTableMarkers()
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
                  <Relationship Id="rSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
                </Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnotePr>
                    <w:numFmt w:val="lowerRoman"/>
                    <w:numStart w:val="4"/>
                  </w:footnotePr>
                  <w:endnotePr>
                    <w:numFmt w:val="upperLetter"/>
                    <w:numStart w:val="2"/>
                  </w:endnotePr>
                </w:settings>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="2"/><w:endnoteReference w:id="3"/></w:r></w:p>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t><w:footnoteReference w:id="4"/></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph bodyParagraph = document.Paragraphs[0];
        TestAssert.Equal("iv", bodyParagraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Footnote).DisplayText ?? string.Empty);
        TestAssert.Equal("B", bodyParagraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Endnote).DisplayText ?? string.Empty);
        DocxParagraph cellParagraph = document.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single();
        TestAssert.Equal("v", cellParagraph.InlineReferences.Single().DisplayText ?? string.Empty);
        TestAssert.Equal("v", cellParagraph.Runs.Last().Text);
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

    public static void DocxRendererEmitsFloatingTextBoxTableRows()
    {
        DocxTable table = DocxTests.CreateTextBoxProbeTable("Text box table cell", "00B0F0");
        DocxFloatingDrawing drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxTableElement(table)]);
        DocxDocument document = new(
            300d,
            300d,
            30d,
            30d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [drawing],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d))],
            [],
            []);
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(1, layout.FloatingDrawings.Single().TextBoxTableRowCount);
        TestAssert.Contains(DocxTests.FormatPdfRgb(((byte)0, (byte)176, (byte)240), "rg"), page.Content);
    }

    public static void DocxRendererEmitsStaticFloatingTextBoxTableRows()
    {
        DocxTable table = DocxTests.CreateTextBoxProbeTable("Static text box table cell", "92D050");
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxTableElement(table)])]
            }
        };
        DocxDocument document = new(
            300d,
            300d,
            30d,
            30d,
            30d,
            30d,
            pageSettings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            []);
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(1, layout.StaticFloatingDrawings.Single().TextBoxTableRowCount);
        TestAssert.Contains(DocxTests.FormatPdfRgb(((byte)146, (byte)208, (byte)80), "rg"), page.Content);
    }

    public static void DocxRendererEmitsPlacedFootnoteFloatingTextBoxTableRows()
    {
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body footnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "51",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 5)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote drawing anchor", 10d, 12d);
        DocxTable table = DocxTests.CreateTextBoxProbeTable("Footnote text box table cell", "FFC000");
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "51",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null)
        {
            FloatingDrawings = [DocxTests.CreateFloatingTextBoxDrawing([new DocxTableElement(table)])]
        };
        DocxDocument document = new(
            300d,
            300d,
            30d,
            30d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            [])
        {
            RelatedStories = [footnoteStory]
        };
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "51");
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(1, placedStory.FloatingDrawings.Single().TextBoxLayout!.TableRows.Count);
        TestAssert.Contains(DocxTests.FormatPdfRgb(((byte)255, (byte)192, (byte)0), "rg"), page.Content);
    }

    public static void DocxSyntheticTableRendersCellsAndText()
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
        TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 1, "Expected DOCX paragraph text to render as a PDF text-show operation.");
    }

    public static void DocxSyntheticTableWithoutBordersDoesNotInventCellGrid()
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
        TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 1, "Expected DOCX table text to render as a PDF text-show operation.");
    }

    public static void DocxSyntheticTableTextRendersWithRunFontResourceWithoutFallback()
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
        TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 1, "Expected DOCX table text to render as a PDF text-show operation.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
                    <w:tcPr><w:noWrap/><w:tcFitText/><w:textDirection w:val="tbRl"/><w:shd w:val="clear" w:color="auto" w:fill="D9EAD3"/></w:tcPr>
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
                        <w:tc><w:tcPr><w:noWrap w:val="0"/><w:tcFitText w:val="0"/><w:textDirection w:val="lrTb"/><w:shd w:fill="FCE5CD"/></w:tcPr><w:p><w:r><w:t>Direct</w:t></w:r></w:p></w:tc>
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal("ShadedTable", document.Tables[0].StyleId ?? string.Empty);
        TestAssert.Equal("CFE2F3", document.Tables[0].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("clear", document.Tables[0].Rows[0].Cells[0].ShadingValue ?? string.Empty);
        TestAssert.Equal("auto", document.Tables[0].Rows[0].Cells[0].ShadingColor ?? string.Empty);
        TestAssert.Equal("top", document.Tables[0].Rows[0].Cells[0].Borders.Single().Edge);
        TestAssert.Equal("16", document.Tables[0].Rows[0].Cells[0].Borders.Single().SizeValue ?? string.Empty);
        TestAssert.Equal("240", document.Tables[0].Rows[0].Cells[0].Margins.LeftValue ?? string.Empty);
        TestAssert.Equal(12d, document.Tables[0].Rows[0].Cells[0].Margins.LeftPoints ?? 0d);
        TestAssert.True(document.Tables[0].Rows[0].Cells[0].NoWrap, "Conditional cell styles should retain base style no-wrap.");
        TestAssert.True(document.Tables[0].Rows[0].Cells[0].FitText, "Conditional cell styles should retain base style fit-text.");
        TestAssert.Equal("tbRl", document.Tables[0].Rows[0].Cells[0].TextDirectionValue ?? string.Empty);
        TestAssert.Equal("FCE5CD", document.Tables[0].Rows[0].Cells[1].FillHex ?? string.Empty);
        TestAssert.True(!document.Tables[0].Rows[0].Cells[1].NoWrap, "Direct no-wrap off should override table style no-wrap.");
        TestAssert.Equal("0", document.Tables[0].Rows[0].Cells[1].NoWrapValue ?? string.Empty);
        TestAssert.True(!document.Tables[0].Rows[0].Cells[1].FitText, "Direct fit-text off should override table style fit-text.");
        TestAssert.Equal("0", document.Tables[0].Rows[0].Cells[1].FitTextValue ?? string.Empty);
        TestAssert.Equal("lrTb", document.Tables[0].Rows[0].Cells[1].TextDirectionValue ?? string.Empty);
        TestAssert.Equal("D9EAD3", document.Tables[0].Rows[1].Cells[0].FillHex ?? string.Empty);
        TestAssert.True(document.Tables[0].Rows[1].Cells[0].NoWrap, "Base table cell style should apply no-wrap to body cells.");
        TestAssert.True(document.Tables[0].Rows[1].Cells[0].FitText, "Base table cell style should apply fit-text to body cells.");
        TestAssert.Equal("tbRl", document.Tables[0].Rows[1].Cells[0].TextDirectionValue ?? string.Empty);
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
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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

        DocxTableStyleDefinitionSummary tableStyle = document.StyleCatalog.TableStyles.Single(style => style.StyleId == "ChildTable");
        TestAssert.True(tableStyle.BasedOnStyleId == "BaseTable" && tableStyle.HasCellProperties && tableStyle.HasParagraphProperties && tableStyle.HasRunProperties && tableStyle.ConditionalRegionCount == 1, "The DOCX style catalog should retain private-safe table style topology after basedOn resolution.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph firstParagraph = document.Tables[0].Rows[0].Cells[0].Paragraphs.Single();
        DocxParagraph secondParagraph = document.Tables[0].Rows[0].Cells[1].Paragraphs.Single();
        TestAssert.True(firstParagraph.StyleResolution.HasTableStyleParagraphProperties, "Table-cell paragraph resolution should record table style participation.");
        TestAssert.True(firstParagraph.StyleResolution.HasDirectParagraphProperties, "Direct spacing in the table-cell paragraph should remain distinguishable from table style properties.");
        TestAssert.True(secondParagraph.StyleResolution.HasTableStyleParagraphProperties, "Inherited table-cell paragraph properties should be visible even without direct pPr.");
        TestAssert.True(secondParagraph.StyleResolution.HasDirectParagraphProperties == false, "A table-cell paragraph without pPr should not report direct paragraph properties.");
        TestAssert.Equal(DocxTextAlignment.Right, firstParagraph.Alignment);
        TestAssert.Equal("right", firstParagraph.AlignmentValue ?? string.Empty);
        TestAssert.True(firstParagraph.EffectiveProperties.StyleResolution.HasTableStyleParagraphProperties, "Effective table-cell paragraph properties should retain table-style provenance.");
        TestAssert.True(firstParagraph.EffectiveProperties.StyleResolution.HasDirectParagraphProperties, "Effective table-cell paragraph properties should retain direct paragraph-property provenance.");
        TestAssert.Equal(DocxTextAlignment.Right, firstParagraph.EffectiveProperties.Alignment);
        TestAssert.Equal(6d, firstParagraph.SpacingAfterPoints);
        TestAssert.Equal(11d, firstParagraph.Runs.Single().FontSize);
        TestAssert.True(firstParagraph.Runs.Single().Italic, "First-column table run style should apply italic.");
        TestAssert.True(firstParagraph.Runs.Single().AllCaps, "First-column table run style should apply all-caps.");
        TestAssert.Equal("FIRST", firstParagraph.Runs.Single().Text);
        TestAssert.Equal("4472C4", firstParagraph.Runs.Single().ColorHex ?? string.Empty);
        TestAssert.Equal(0d, secondParagraph.SpacingAfterPoints);
        TestAssert.True(secondParagraph.EffectiveProperties.StyleResolution.HasTableStyleParagraphProperties, "Effective inherited table-cell properties should remain tied to the table style.");
        TestAssert.True(secondParagraph.EffectiveProperties.StyleResolution.HasDirectParagraphProperties == false, "Effective inherited table-cell properties should not pretend to be direct overrides.");
        TestAssert.Equal(0d, secondParagraph.EffectiveProperties.SpacingAfterPoints);
        TestAssert.Equal(11d, secondParagraph.Runs.Single().FontSize);
    }

    public static void DocxTableStyleParagraphSpacingCollapsesInCellLayout()
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
                  <w:style w:type="table" w:styleId="CellText">
                    <w:pPr>
                      <w:spacing w:before="240" w:after="120"/>
                      <w:contextualSpacing/>
                    </w:pPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="CellText"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:p><w:r><w:t>First styled cell paragraph</w:t></w:r></w:p>
                          <w:p><w:r><w:t>Second styled cell paragraph</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph[] paragraphs = document.Tables.Single().Rows.Single().Cells.Single().Paragraphs.ToArray();
        TestAssert.True(paragraphs.All(paragraph => paragraph.EffectiveProperties.StyleResolution.HasTableStyleParagraphProperties), "Table style paragraph spacing should remain visible in the effective cell paragraph model.");

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single()
            .Cells.Single()
            .TextLines.ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(12d, lines[0].ParagraphBeforeSpacing ?? -1d);
        TestAssert.Equal(12d, lines[0].AppliedBeforeSpacing ?? -1d);
        TestAssert.Equal(6d, lines[0].ParagraphAfterSpacing ?? -1d);
        TestAssert.Equal(6d, lines[1].PendingAfterSpacing ?? -1d);
        TestAssert.Equal(12d, lines[1].ParagraphBeforeSpacing ?? -1d);
        TestAssert.Equal(0d, lines[1].AppliedBeforeSpacing ?? -1d);
        TestAssert.True(lines[1].ContextualSpacingSuppressed == true, "Same-style table-cell paragraphs should collapse table-style contextual spacing before layout.");
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

        DocxDocument document = new DocxReader().Read(OoxPackage.Open(input, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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

        DocxDocument document = new DocxReader().Read(OoxPackage.Open(input, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxTextRun run = document.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single().Runs.Single();

        TestAssert.Equal(14d, run.FontSize);
        TestAssert.Equal("CC0000", run.ColorHex ?? string.Empty);
        TestAssert.True(run.Bold, "Paragraph-style run properties should remain above table-style run properties.");
        TestAssert.True(run.Italic, "Non-conflicting table-style italic should still apply.");
        TestAssert.True(run.AllCaps, "Non-conflicting table-style caps should still apply.");
        TestAssert.Equal("CONFLICT", run.Text);
        TestAssert.Equal("StrongRed", run.StyleResolution.CharacterStyleId ?? string.Empty);
        TestAssert.True(run.StyleResolution.HasTableStyleRunProperties, "Table-cell run provenance should retain table-style contribution.");
        TestAssert.True(run.StyleResolution.HasParagraphStyleRunProperties, "Table-cell run provenance should retain paragraph-style run contribution.");
        TestAssert.True(run.StyleResolution.HasCharacterStyleRunProperties, "Table-cell run provenance should retain character-style run contribution.");
        TestAssert.True(run.StyleResolution.HasDirectRunProperties == false, "An rPr containing only rStyle should not be treated as a direct run override.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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

    public static void DocxReaderTableStyleBandsLeadingEdgeWithoutHeaderEmphasis()
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
                  <w:style w:type="table" w:styleId="LeadBand">
                    <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/></w:tblPr>
                    <w:tcPr><w:shd w:fill="DDDDDD"/></w:tcPr>
                    <w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="firstCol"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="lastRow"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="lastCol"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="00FF00"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="band1Vert"><w:tcPr><w:shd w:fill="00FF00"/></w:tcPr></w:tblStylePr>
                    <w:tblStylePr w:type="band2Vert"><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr></w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="LeadBand"/><w:tblLook w:val="0000" w:firstRow="0" w:firstColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>A0</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="LeadBand"/><w:tblLook w:val="0000" w:firstRow="0" w:firstColumn="0" w:noHBand="1" w:noVBand="0"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>B0</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="LeadBand"/><w:tblLook w:val="0020" w:firstRow="1" w:firstColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>C0</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="LeadBand"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>D00</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>D01</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>D10</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>D11</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="LeadBand"/><w:tblLook w:val="00A0" w:firstRow="1" w:firstColumn="1" w:noHBand="0" w:noVBand="0"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>E00</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>E01</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>E10</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>E11</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>E20</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>E21</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="LeadBand"/><w:tblLook w:val="0140" w:firstRow="0" w:lastRow="1" w:firstColumn="0" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>F00</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>F01</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>F10</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>F11</w:t></w:r></w:p></w:tc>
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

        TestAssert.Equal("00FF00", document.Tables[0].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("0000FF", document.Tables[0].Rows[1].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("00FF00", document.Tables[0].Rows[2].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("00FF00", document.Tables[1].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("0000FF", document.Tables[1].Rows[0].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("00FF00", document.Tables[1].Rows[0].Cells[2].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[2].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("00FF00", document.Tables[2].Rows[1].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[3].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[3].Rows[0].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[3].Rows[1].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("00FF00", document.Tables[3].Rows[1].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[4].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[4].Rows[0].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[4].Rows[1].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("00FF00", document.Tables[4].Rows[1].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[4].Rows[2].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("0000FF", document.Tables[4].Rows[2].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("00FF00", document.Tables[5].Rows[0].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[5].Rows[0].Cells[1].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[5].Rows[1].Cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("FF0000", document.Tables[5].Rows[1].Cells[1].FillHex ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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

    public static void DocxSyntheticCommonTableBorderStylesRenderAndInspect()
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
                              <w:top w:val="double" w:color="FF0000" w:sz="18"/>
                              <w:bottom w:val="single" w:color="111111" w:sz="8"/>
                              <w:left w:val="dotted" w:color="00AA00" w:sz="8"/>
                              <w:right w:val="dashed" w:color="0000FF" w:sz="8"/>
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
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0 0.667 0 rg", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 7, "Common border styles should render as filled strip geometry, including segmented and double-strip variants.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(4, table.VisibleBorderCount);
        TestAssert.Equal(1, table.SingleBorderCount);
        TestAssert.Equal(1, table.DoubleBorderCount);
        TestAssert.Equal(1, table.DottedBorderCount);
        TestAssert.Equal(1, table.DashedBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticDotDashTableBorderStylesEmitPatternedSegments()
    {
        int dashedRectangles = RenderBorderRectangleCount("dashed");
        int dotDashRectangles = RenderBorderRectangleCount("dotDash");
        int dotDotDashRectangles = RenderBorderRectangleCount("dotDotDash");

        TestAssert.True(dotDashRectangles > dashedRectangles, "dotDash should emit dash and dot segments, not plain dash repetition.");
        TestAssert.True(dotDotDashRectangles > dotDashRectangles, "dotDotDash should emit an extra dot segment in each pattern cycle.");

        static int RenderBorderRectangleCount(string borderStyle)
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
                ["word/document.xml"] = $"""
                    <?xml version="1.0" encoding="UTF-8"?>
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:body>
                        <w:tbl>
                          <w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid>
                          <w:tr>
                            <w:tc>
                              <w:tcPr>
                                <w:tcBorders>
                                  <w:top w:val="{borderStyle}" w:color="336699" w:sz="8"/>
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
            return pdf.Split(" re f", StringSplitOptions.None).Length - 1;
        }
    }

    public static void DocxSyntheticDashDotStrokedTableBorderStyleRendersWithoutDiagnostic()
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
                      <w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="dashDotStroked" w:color="336699" w:sz="12"/>
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
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 3, "dashDotStroked borders should render dash-dot segments plus a center stroke.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(1, table.VisibleBorderCount);
        TestAssert.Equal(1, table.DashedBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticOutsetTableBorderStyleRendersWithoutDiagnostic()
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
                              <w:top w:val="outset" w:color="336699" w:sz="12"/>
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
        TestAssert.Contains("0.2 0.4 0.6 rg", pdf);
        TestAssert.Contains(" re f", pdf);

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(1, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticInsetTableBorderStyleRendersWithoutDiagnostic()
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
                              <w:top w:val="inset" w:color="663399" w:sz="12"/>
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
        TestAssert.Contains("0.4 0.2 0.6 rg", pdf);
        TestAssert.Contains(" re f", pdf);

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(1, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxTableBottomBorderHangsBelowLastRowContent()
    {
        // Office A/B (w7 doc-start table probe, Word-COM rendered plus PdfInspect border
        // rects): Word hangs the last row bottom border BELOW cell content (outside:
        // 689.14 to 689.62 under content ending 689.62), while shared interior borders
        // straddle boundaries pitch-neutrally. The table terminus must consume the bottom
        // width, or following flow rides high (here by a full border width).
        static DocxLayout LayoutPair(bool withBottomBorder, out double tableTotal, out double afterBaseline)
        {
            var first = new DocxTableCell("A1", [DocxTests.CreateDocxLayoutParagraph("A1", 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty);
            IReadOnlyList<DocxTableCellBorder> bottom = withBottomBorder
                ? [new DocxTableCellBorder("bottom", "single", "auto", "8")]
                : [];
            var last = new DocxTableCell("B1", [DocxTests.CreateDocxLayoutParagraph("B1", 10d, 10d)], null, null, null, null, bottom, DocxTableCellMargins.Empty);
            var table = new DocxTable(null, [180d], [new DocxTableRow([first], null), new DocxTableRow([last], null)]);
            DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
            DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxTableElement(table), new DocxParagraphElement(after)],
                [table]);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            DocxTableRowLayout[] rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            tableTotal = rows.Sum(row => row.Height);
            afterBaseline = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single(line => line.Text == "After").BaselineY;
            return layout;
        }

        _ = LayoutPair(false, out double plainTotal, out double plainAfter);
        _ = LayoutPair(true, out double borderedTotal, out double borderedAfter);
        // sz=8 resolves to a 0.96pt visible width; the bordered table carries one 0.96
        // collapsed advance plus one 0.96 terminus extension over the plain table.
        TestAssert.True(Math.Abs((borderedTotal - plainTotal) - 1.92d) < 0.000001d, "Bordered table should exceed the plain table by advance plus terminus. Delta=" + (borderedTotal - plainTotal).ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs((plainAfter - borderedAfter) - 1.92d) < 0.000001d, "Following text should ride lower by advance plus terminus. Shift=" + (plainAfter - borderedAfter).ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxTableExactLastRowBottomBorderHangsBelowContent()
    {
        // Office A/B (w8 exact-row probe, Word-COM rendered plus PdfInspect border
        // rects): an exact-36 bordered row renders 36.5 tall with content top-anchored
        // exactly like atLeast, so the exact height pins the content box and the
        // bottom width still hangs below it (sz=8 resolves to 0.96).
        var paragraph = DocxTests.CreateDocxLayoutParagraph("A", 10d, 10d);
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, [new DocxTableCellBorder("bottom", "single", "auto", "8")], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 10d) with {HeightValue = "200", HeightRuleValue = "exact" }]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(Math.Abs(row.Height - 10.96d) < 0.000001d, "Exact last row should hang the bottom width below the declared height. Height=" + row.Height.ToString(CultureInfo.InvariantCulture));
    }

    public static void RunVerticalAlignmentParsesKnownSpellings()
    {
        // R17: superscript/subscript parse case-insensitively; everything else falls
        // back to Baseline like the legacy comparisons, which only special-cased the
        // two script shifts.
        TestAssert.Equal(DocxRunVerticalAlignment.Superscript, DocxTextRun.ParseVerticalAlignment("superscript"));
        TestAssert.Equal(DocxRunVerticalAlignment.Superscript, DocxTextRun.ParseVerticalAlignment("SuperScript"));
        TestAssert.Equal(DocxRunVerticalAlignment.Subscript, DocxTextRun.ParseVerticalAlignment("subscript"));
        TestAssert.Equal(DocxRunVerticalAlignment.Subscript, DocxTextRun.ParseVerticalAlignment("SUBSCRIPT"));
        TestAssert.Equal(DocxRunVerticalAlignment.Baseline, DocxTextRun.ParseVerticalAlignment("baseline"));
        TestAssert.Equal(DocxRunVerticalAlignment.Baseline, DocxTextRun.ParseVerticalAlignment(string.Empty));
        TestAssert.Equal(DocxRunVerticalAlignment.Baseline, DocxTextRun.ParseVerticalAlignment(null));
        TestAssert.Equal(DocxRunVerticalAlignment.Baseline, DocxTextRun.ParseVerticalAlignment("raised"));
    }

    public static void TableWidthKindParsesKnownSpellings()
    {
        // R17: dxa/pct parse case-insensitively; everything else falls back to Auto
        // like the legacy comparisons, which only special-cased the two kinds.
        TestAssert.Equal(DocxTableWidthKind.Dxa, DocxTable.ParseTableWidthKind("dxa"));
        TestAssert.Equal(DocxTableWidthKind.Dxa, DocxTable.ParseTableWidthKind("DXA"));
        TestAssert.Equal(DocxTableWidthKind.Percent, DocxTable.ParseTableWidthKind("pct"));
        TestAssert.Equal(DocxTableWidthKind.Percent, DocxTable.ParseTableWidthKind("Pct"));
        TestAssert.Equal(DocxTableWidthKind.Auto, DocxTable.ParseTableWidthKind("auto"));
        TestAssert.Equal(DocxTableWidthKind.Auto, DocxTable.ParseTableWidthKind(string.Empty));
        TestAssert.Equal(DocxTableWidthKind.Auto, DocxTable.ParseTableWidthKind(null));
        TestAssert.Equal(DocxTableWidthKind.Auto, DocxTable.ParseTableWidthKind("nimble"));
    }

    public static void CellVerticalAlignmentParsesKnownSpellings()
    {
        // R17: bottom/center parse case-insensitively; everything else (including
        // "both", empty, and unknown spellings) falls back to Top like the legacy
        // comparisons, which only special-cased bottom and center.
        TestAssert.Equal(DocxTableCellVerticalAlignment.Bottom, DocxTableCell.ParseVerticalAlignment("bottom"));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Bottom, DocxTableCell.ParseVerticalAlignment("BOTTOM"));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Center, DocxTableCell.ParseVerticalAlignment("center"));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Center, DocxTableCell.ParseVerticalAlignment("Center"));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Top, DocxTableCell.ParseVerticalAlignment("top"));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Top, DocxTableCell.ParseVerticalAlignment("both"));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Top, DocxTableCell.ParseVerticalAlignment(string.Empty));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Top, DocxTableCell.ParseVerticalAlignment(null));
        TestAssert.Equal(DocxTableCellVerticalAlignment.Top, DocxTableCell.ParseVerticalAlignment("justified"));
    }
    // RV06 (Office gate): autofit tables distribute column widths by content instead
    // of the declared grid (fails: second column starts at the grid position).
    public static void DocxAutofitTableDistributesWidthsByContent()
    {
        string input = FindCase("docx-markup-review.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });
        string pdf = File.ReadAllText(output, Encoding.Latin1);
        if (!pdf.Contains("Aptos", StringComparison.Ordinal))
        {
            TestAssert.Skip("Environmental precondition not met: (Aptos not embedded)");
        }

        // Header-row cell clips: the first column anchors at the table edge while
        // Office autofit moves the second column to its measured split.
        var rows = new List<(double Y, double X)>();
        foreach (Match match in Regex.Matches(pdf, @"(\d+\.?\d*) (\d+\.?\d*) \d+\.?\d* \d+\.?\d* re W n"))
        {
            rows.Add((
                double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)));
        }

        TestAssert.True(rows.Count != 0, "Expected table cell clips.");
        double headerY = rows.Max(row => row.Y);
        double[] xs = rows.Where(row => Math.Abs(row.Y - headerY) < 2d).Select(row => row.X).Distinct().OrderBy(x => x).ToArray();
        TestAssert.Equal(2, xs.Length);
        TestAssert.True(Math.Abs(xs[0] - 72d) < 2d, "The first column must anchor at the table edge.");
        TestAssert.True(xs[1] > 185d && xs[1] < 196d, "Autofit must split columns at the Office-measured position.");
    }

    // RV06: table autofit measures deleted text, so the reader preserves Final-view
    // excluded del runs for measurement (rendering keeps filtering them).
    public static void DocxReaderPreservesDeletedTextForMeasurement()
    {
        string input = FindCase("docx-markup-review.docx");
        using FileStream stream = File.OpenRead(input);
        DocxDocument final = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxTable finalTable = final.BodyElements.OfType<DocxTableElement>().First().Table;
        DocxParagraph finalCell = ((DocxParagraphElement)finalTable.Rows[1].Cells[1].BodyElements[0]).Paragraph;
        TestAssert.Equal("removed cell text", finalCell.DeletedText);
        using FileStream originalStream = File.OpenRead(input);
        DocxDocument original = new DocxReader().Read(OoxPackage.Open(originalStream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Original);
        DocxTable originalTable = original.BodyElements.OfType<DocxTableElement>().First().Table;
        DocxParagraph originalCell = ((DocxParagraphElement)originalTable.Rows[1].Cells[1].BodyElements[0]).Paragraph;
        TestAssert.Equal("added cell text", originalCell.DeletedText);
    }

    // RV06 (Office gate): autofit ignores grid variation, so a differentiated grid
    // with balanced content still splits by content (fails: second column at grid).
    public static void DocxAutofitTableIgnoresGridVariation()
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
                        <w:tblW w:w="7200" w:type="dxa"/>
                        <w:tblBorders>
                          <w:top w:val="single" w:sz="6" w:color="808080"/>
                          <w:left w:val="single" w:sz="6" w:color="808080"/>
                          <w:bottom w:val="single" w:sz="6" w:color="808080"/>
                          <w:right w:val="single" w:sz="6" w:color="808080"/>
                          <w:insideH w:val="single" w:sz="6" w:color="808080"/>
                          <w:insideV w:val="single" w:sz="6" w:color="808080"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="6000"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Alpha Beta</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Gamma Delta</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1440"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });
        string pdf = File.ReadAllText(output, Encoding.Latin1);
        if (!pdf.Contains("Aptos", StringComparison.Ordinal))
        {
            TestAssert.Skip("Environmental precondition not met: (Aptos not embedded)");
        }

        var starts = new List<double>();
        foreach (Match match in Regex.Matches(pdf, @"(\d+\.?\d*) \d+\.?\d* \d+\.?\d* \d+\.?\d* re W n"))
        {
            starts.Add(double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture));
        }

        double[] xs = starts.Distinct().OrderBy(x => x).ToArray();
        TestAssert.Equal(2, xs.Length);
        TestAssert.True(xs[1] > 222d && xs[1] < 239d, "Autofit must split balanced content at the Office-measured position, not the grid.");
    }

    private static string FindCase(string name)
    {
        string[] candidates = new[]
        {
            Path.Combine("tests", "Lokad.OoxPdf.Tests", "Cases", name),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name),
        };
        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException("Case file not found: " + name);
    }
}

