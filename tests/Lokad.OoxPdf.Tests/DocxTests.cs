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
    internal static void AssertSimpleMarkupTableRowChangeBarRendered(DocxDocument document, DocxTable table, string flowName)
    {
        string content = string.Concat(new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Select(page => page.Content));
        string color = FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot(table.Rows.Single().Revisions), "rg");
        TestAssert.True(content.Contains(color, StringComparison.Ordinal), "Simple-markup should draw a row revision color for " + flowName + ".");
        TestAssert.True(content.Contains(" 1.5 ", StringComparison.Ordinal), "Simple-markup should draw a narrow row change bar for " + flowName + ".");
    }

    internal static DocxDocument CreateCommentMarkerFlowDocument(
        IReadOnlyList<DocxBodyElement> bodyElements,
        IReadOnlyList<DocxFloatingDrawing> floatingDrawings,
        DocxPageSettings settings,
        IReadOnlyList<DocxTable> fallbackTables)
    {
        return new DocxDocument(
            300d,
            300d,
            30d,
            30d,
            45d,
            30d,
            settings,
            floatingDrawings,
            [],
            [],
            bodyElements,
            [],
            fallbackTables)
        {
            MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
        };
    }

    internal static void AssertSimpleMarkupCommentMarkerRendered(DocxDocument document, string flowName)
    {
        DocxLayoutSnapshot layout = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);
        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(CountOccurrences(page.Content, "1 0.753 0 rg") >= 1, "Comment markers should use the expected marker fill color for " + flowName + ".");
        TestAssert.True(CountOccurrences(page.Content, "0.851 0.592 0 RG") >= 1, "Comment markers should use the expected marker outline color for " + flowName + ".");
        TestAssert.True(CountOccurrences(page.Content, " re f") >= 1, "Simple-markup rendering should emit a filled comment marker rectangle for " + flowName + ".");
        TestAssert.True(CountOccurrences(page.Content, " re S") >= 1, "Simple-markup rendering should emit a stroked comment marker outline for " + flowName + ".");
        TestAssert.True(
            layout.Pages[0].TextLineCount +
            layout.Pages[0].StaticTextLineCount +
            layout.Pages[0].StaticTableRowCount +
            layout.Pages[0].PlacedRelatedStoryTextLineCount +
            layout.Pages[0].PlacedRelatedStoryTableRowCount +
            layout.Pages[0].TableRows.Sum(row => row.TextLineCount) +
            layout.FloatingDrawings.Sum(drawing => drawing.TextBoxTextLineCount) +
            layout.FloatingDrawings.Sum(drawing => drawing.TextBoxTableRowCount) +
            layout.StaticFloatingDrawings.Sum(drawing => drawing.TextBoxTableRowCount) +
            layout.StaticFloatingDrawings.Sum(drawing => drawing.TextBoxTextLineCount) > 0,
            "The " + flowName + " fixture should include a laid-out text line.");
    }

    internal static void AssertWordCompatibleCommentRangeMarkerRendered(DocxDocument document, string flowName)
    {
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(
            CountOccurrences(page.Content, " 11.625 re f") >= 1,
            "Word-compatible all-markup should emit comment range fill geometry for " + flowName + ".");
        TestAssert.Contains("0.82 0.204 0.22 RG", page.Content);
        TestAssert.True(
            layout.Pages[0].TextLineCount +
            layout.Pages[0].StaticTextLineCount +
            layout.Pages[0].StaticTableRowCount +
            layout.Pages[0].PlacedRelatedStoryTextLineCount +
            layout.Pages[0].PlacedRelatedStoryTableRowCount +
            layout.Pages[0].TableRows.Sum(row => row.TextLineCount) +
            layout.FloatingDrawings.Sum(drawing => drawing.TextBoxTextLineCount) +
            layout.FloatingDrawings.Sum(drawing => drawing.TextBoxTableRowCount) +
            layout.StaticFloatingDrawings.Sum(drawing => drawing.TextBoxTableRowCount) +
            layout.StaticFloatingDrawings.Sum(drawing => drawing.TextBoxTextLineCount) > 0,
            "The " + flowName + " fixture should include a laid-out text line.");
    }

    internal static DocxTable CreateCommentRangeTable(string text, string commentId)
    {
        DocxParagraph paragraph = CreateCommentRangeParagraph(text, commentId);
        var cell = new DocxTableCell(
            text,
            [paragraph],
            FillHex: null,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders: [],
            DocxTableCellMargins.Empty);
        return new DocxTable(
            LayoutValue: null,
            ColumnWidthsPoints: [100d],
            Rows: [new DocxTableRow([cell], HeightPoints: 18d)]);
    }

    internal static DocxParagraph CreateCommentRangeParagraph(string text, string commentId)
    {
        DocxTextRun run = new(text, 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        return new DocxParagraph(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Comment,
                    commentId,
                    CustomMarkFollowsValue: null,
                    SourceRunIndex: 1,
                    RunChildIndex: 0,
                    TextOffsetInRun: 0, DisplayText: null)
            ],
            CommentRanges =
            [
                new DocxCommentRange(commentId, 0, 0, 1, text.Length, 1, 0)
            ]
        };
    }

    internal static DocxParagraph CreateCommentMarkerParagraph(string text, string commentId)
    {
        return CreateDocxLayoutParagraph(text, 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Comment,
                    commentId,
                    CustomMarkFollowsValue: null,
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 0, DisplayText: null)
            ]
        };
    }

    internal static DocxFloatingDrawing CreateFloatingTextBoxDrawing(IReadOnlyList<DocxBodyElement> textBoxBodyElements)
    {
        return new DocxFloatingDrawing(
            DistanceTopValue: "0",
            DistanceBottomValue: "0",
            DistanceLeftValue: "0",
            DistanceRightValue: "0",
            SimplePositionValue: "0",
            RelativeHeightValue: "0",
            BehindDocumentValue: "0",
            LockedValue: "0",
            LayoutInCellValue: "1",
            AllowOverlapValue: "1",
            ExtentCxValue: "1371600",
            ExtentCyValue: "457200",
            HorizontalRelativeFromValue: "page",
            HorizontalAlignValue: null,
            HorizontalOffsetValue: "914400",
            VerticalRelativeFromValue: "page",
            VerticalAlignValue: null,
            VerticalOffsetValue: "914400",
            WrapKind: DocxFloatingWrapKind.None,
            WrapTextValue: null,
            SourceParagraphIndex: 0,
            SourceBlockIndex: 0, ImageRelationshipId: null, Image: null)
        {
            TextBoxBodyElements = textBoxBodyElements
        };
    }

    internal static DocxTable CreateTextBoxProbeTable(string text, string fillHex)
    {
        DocxParagraph cellParagraph = CreateDocxLayoutParagraph(text, 10d, 12d);
        var cell = new DocxTableCell(
            text,
            [cellParagraph],
            FillHex: fillHex,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders: [],
            DocxTableCellMargins.Empty);
        return new DocxTable(
            LayoutValue: null,
            ColumnWidthsPoints: [100d],
            Rows: [new DocxTableRow([cell], HeightPoints: 24d)]);
    }

    internal static DocxTable CreateTableRevisionBalloonProbeTable(string text, string idPrefix)
    {
        DocxTableCell cell = new(text, [CreateDocxLayoutParagraph(text, 10d, 12d)], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            Revisions =
            [
                new DocxRevisionInfo(
                    DocxRevisionKind.TableCellPropertiesChange,
                    idPrefix + "c",
                    "Reviewer",
                    "2026-06-10T00:00:00Z",
                    "tcPrChange",
                    DocxRevisionPropertyFamily.Cell,
                    ["tcW"])
            ]
        };
        DocxTableRow row = new([cell], 18d)
        {
            Revisions =
            [
                new DocxRevisionInfo(
                    DocxRevisionKind.TableRowPropertiesChange,
                    idPrefix + "r",
                    "Reviewer",
                    "2026-06-10T00:00:00Z",
                    "trPrChange",
                    DocxRevisionPropertyFamily.Row,
                    ["trHeight"])
            ]
        };
        return new DocxTable(null, [90d], [row])
        {
            Revisions =
            [
                new DocxRevisionInfo(
                    DocxRevisionKind.TablePropertiesChange,
                    idPrefix + "t",
                    "Reviewer",
                    "2026-06-10T00:00:00Z",
                    "tblPrChange",
                    DocxRevisionPropertyFamily.Table,
                    ["tblBorders"])
            ]
        };
    }

    internal static int CountRevisionBalloonCandidates(IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements)
    {
        return placements
            .Where(placement => placement.Kind == "Revision")
            .Sum(placement => placement.CandidateCount);
    }

    internal static string FormatPdfRgb((byte Red, byte Green, byte Blue) color, string operatorName)
    {
        return FormatPdfColor(color.Red) + " " + FormatPdfColor(color.Green) + " " + FormatPdfColor(color.Blue) + " " + operatorName;
    }

    private static string FormatPdfColor(byte value)
    {
        return (value / 255d).ToString("0.###", CultureInfo.InvariantCulture);
    }

    internal static DocxParagraph ReadTrackedChangeModeProbe(string input, OoxPdfDocxMarkupMode mode)
    {
        return ReadDocx(input, mode).Paragraphs.Single();
    }

    internal static DocxDocument ReadDocx(string input, OoxPdfDocxMarkupMode mode)
    {
        using FileStream stream = File.OpenRead(input);
        return new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: mode);
    }

    internal static DocxTable ReadSingleTable(string input, OoxPdfDocxMarkupMode mode)
    {
        return ReadDocx(input, mode).BodyElements.OfType<DocxTableElement>().Single().Table;
    }

    internal static string ParagraphTexts(DocxDocument document)
    {
        return string.Join("|", document.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
    }

    internal static string RowTexts(DocxTable table)
    {
        return string.Join("|", table.Rows.Select(row => string.Join(" ", row.Cells.Select(cell => cell.Text))));
    }

    internal static string WriteTrackedChangeModeProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                      <w:ins w:id="1" w:author="A" w:date="2026-06-01T00:00:00Z"><w:r><w:t>Inserted </w:t></w:r></w:ins>
                      <w:del w:id="2" w:author="B" w:date="2026-06-02T00:00:00Z"><w:r><w:delText>Deleted </w:delText></w:r></w:del>
                      <w:moveFrom w:id="3" w:author="C" w:date="2026-06-03T00:00:00Z"><w:r><w:delText>MovedFrom </w:delText></w:r></w:moveFrom>
                      <w:moveTo w:id="4" w:author="D" w:date="2026-06-04T00:00:00Z"><w:r><w:t>MovedTo </w:t></w:r></w:moveTo>
                      <w:r><w:t>After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteTrackedChangeRevisionViewProbeDocx(string revisionViewAttributes)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:ins w:id="1" w:author="A" w:date="2026-06-01T00:00:00Z"><w:r><w:t>Inserted </w:t></w:r></w:ins>
                      <w:del w:id="2" w:author="B" w:date="2026-06-02T00:00:00Z"><w:r><w:delText>Deleted </w:delText></w:r></w:del>
                      <w:moveFrom w:id="3" w:author="C" w:date="2026-06-03T00:00:00Z"><w:r><w:delText>MovedFrom </w:delText></w:r></w:moveFrom>
                      <w:moveTo w:id="4" w:author="D" w:date="2026-06-04T00:00:00Z"><w:r><w:t>MovedTo </w:t></w:r></w:moveTo>
                      <w:r><w:t>After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/settings.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:revisionView {revisionViewAttributes}/>
                </w:settings>
                """
        });
    }

    internal static string WriteBlockRevisionParagraphProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                    <w:ins w:id="31" w:author="A" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:t>Inserted paragraph</w:t></w:r></w:p>
                    </w:ins>
                    <w:del w:id="32" w:author="B" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:delText>Deleted paragraph</w:delText></w:r></w:p>
                    </w:del>
                    <w:moveFrom w:id="33" w:author="C" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:delText>Moved from paragraph</w:delText></w:r></w:p>
                    </w:moveFrom>
                    <w:moveTo w:id="34" w:author="D" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:t>Moved to paragraph</w:t></w:r></w:p>
                    </w:moveTo>
                    <w:p><w:r><w:t>After</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteBlockRevisionParagraphPropertiesProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                    <w:ins w:id="611" w:author="A" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr>
                          <w:pageBreakBefore/>
                          <w:keepNext/>
                          <w:keepLines/>
                          <w:widowControl w:val="0"/>
                        </w:pPr>
                        <w:r><w:t>Inserted keep</w:t></w:r>
                      </w:p>
                    </w:ins>
                    <w:del w:id="612" w:author="B" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr>
                          <w:pageBreakBefore/>
                          <w:keepNext/>
                          <w:keepLines/>
                          <w:widowControl w:val="0"/>
                        </w:pPr>
                        <w:r><w:delText>Deleted keep</w:delText></w:r>
                      </w:p>
                    </w:del>
                    <w:moveFrom w:id="613" w:author="C" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr>
                          <w:keepNext/>
                          <w:keepLines/>
                          <w:widowControl/>
                        </w:pPr>
                        <w:r><w:delText>Moved from keep</w:delText></w:r>
                      </w:p>
                    </w:moveFrom>
                    <w:moveTo w:id="614" w:author="D" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr>
                          <w:keepNext/>
                          <w:keepLines/>
                          <w:widowControl/>
                        </w:pPr>
                        <w:r><w:t>Moved to keep</w:t></w:r>
                      </w:p>
                    </w:moveTo>
                    <w:p><w:r><w:t>After</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteRevisedParagraphSpacingCollapseProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                  <w:style w:type="paragraph" w:styleId="BodySpacing"/>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="BodySpacing"/>
                        <w:spacing w:after="480"/>
                      </w:pPr>
                      <w:r><w:t>Before</w:t></w:r>
                    </w:p>
                    <w:ins w:id="621" w:author="A" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val="BodySpacing"/>
                          <w:spacing w:before="480" w:after="120"/>
                          <w:contextualSpacing/>
                        </w:pPr>
                        <w:r><w:t>Inserted spacing</w:t></w:r>
                      </w:p>
                    </w:ins>
                    <w:del w:id="622" w:author="B" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val="BodySpacing"/>
                          <w:spacing w:before="480" w:after="120"/>
                          <w:contextualSpacing/>
                        </w:pPr>
                        <w:r><w:delText>Deleted spacing</w:delText></w:r>
                      </w:p>
                    </w:del>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="BodySpacing"/>
                      </w:pPr>
                      <w:r><w:t>After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteDeletedContentProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                        <w:rPr>
                          <w:del w:id="202" w:author="B" w:date="2026-06-10T00:00:00Z"/>
                        </w:rPr>
                      </w:pPr>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:del w:id="201" w:author="A" w:date="2026-06-10T00:00:00Z">
                        <w:r><w:t>deleted-run </w:t></w:r>
                      </w:del>
                      <w:r><w:t>After</w:t></w:r>
                    </w:p>
                    <w:del w:id="203" w:author="C" w:date="2026-06-10T00:00:00Z">
                      <w:tbl>
                        <w:tr><w:tc><w:p><w:r><w:t>Deleted table text</w:t></w:r></w:p></w:tc></w:tr>
                      </w:tbl>
                    </w:del>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteDeletedParagraphMarkMarkupProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/deleted-paragraph-mark" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:spacing w:after="480"/>
                        <w:rPr><w:del w:id="501" w:author="A" w:date="2026-06-10T00:00:00Z"/></w:rPr>
                      </w:pPr>
                      <w:r><w:t>First</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:pPr><w:spacing w:before="480" w:after="120"/></w:pPr>
                      <w:r><w:t>Second </w:t></w:r>
                      <w:hyperlink r:id="rIdLink"><w:r><w:t>link </w:t></w:r></w:hyperlink>
                      <w:fldSimple w:instr=" REF SyntheticTarget "><w:r><w:t>field </w:t></w:r></w:fldSimple>
                      <w:r><w:t>[c]</w:t><w:commentReference w:id="1"/></w:r>
                    </w:p>
                    <w:del w:id="502" w:author="B" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
                        <w:r><w:delText>Deleted list item</w:delText></w:r>
                      </w:p>
                    </w:del>
                    <w:p>
                      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
                      <w:r><w:t>Visible list item</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="1">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
                </w:numbering>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="1"><w:p><w:r><w:t>Synthetic comment</w:t></w:r></w:p></w:comment>
                </w:comments>
                """
        });
    }

    internal static string WriteHyperlinkFieldRevisionProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/revisions" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:hyperlink r:id="rIdLink">
                        <w:r><w:t>Link </w:t></w:r>
                        <w:del w:id="301" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>old </w:delText></w:r></w:del>
                        <w:ins w:id="302" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:t>new </w:t></w:r></w:ins>
                      </w:hyperlink>
                      <w:fldSimple w:instr=" DATE ">
                        <w:del w:id="303" w:author="C" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>field-old </w:delText></w:r></w:del>
                        <w:ins w:id="304" w:author="D" w:date="2026-06-10T00:00:00Z"><w:r><w:t>field-new </w:t></w:r></w:ins>
                      </w:fldSimple>
                      <w:r><w:t>After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteSourceIndexMarkupProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/source-index" TargetMode="External"/>
                  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:del w:id="401" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>Hidden </w:delText></w:r></w:del>
                      <w:ins w:id="402" w:author="A" w:date="2026-06-10T00:00:00Z">
                        <w:bookmarkStart w:id="9" w:name="Target"/>
                        <w:r><w:t>Target </w:t></w:r>
                        <w:bookmarkEnd w:id="9"/>
                      </w:ins>
                      <w:hyperlink w:anchor="Target"><w:r><w:t>Jump </w:t></w:r></w:hyperlink>
                      <w:hyperlink r:id="rIdLink">
                        <w:del w:id="403" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>Old link </w:delText></w:r></w:del>
                        <w:ins w:id="404" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>External </w:t></w:r></w:ins>
                      </w:hyperlink>
                      <w:ins w:id="405" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>Foot</w:t><w:footnoteReference w:id="2"/></w:r></w:ins>
                      <w:r><w:t> End</w:t><w:endnoteReference w:id="3"/></w:r>
                    </w:p>
                    <w:tbl>
                      <w:del w:id="406" w:author="A" w:date="2026-06-10T00:00:00Z">
                        <w:tr><w:tc><w:p><w:r><w:delText>Deleted row</w:delText></w:r></w:p></w:tc></w:tr>
                      </w:del>
                      <w:tr><w:tc><w:p><w:r><w:t>Visible row</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
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
    }

    internal static string WriteBlockRevisionTableProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                      <w:tr><w:tc><w:p><w:r><w:t>Base</w:t></w:r></w:p></w:tc></w:tr>
                      <w:ins w:id="41" w:author="A" w:date="2026-06-10T00:00:00Z">
                        <w:tr><w:tc><w:p><w:r><w:t>Inserted row</w:t></w:r></w:p></w:tc></w:tr>
                      </w:ins>
                      <w:del w:id="42" w:author="B" w:date="2026-06-10T00:00:00Z">
                        <w:tr><w:tc><w:p><w:r><w:delText>Deleted row</w:delText></w:r></w:p></w:tc></w:tr>
                      </w:del>
                      <w:moveFrom w:id="43" w:author="C" w:date="2026-06-10T00:00:00Z">
                        <w:tr><w:tc><w:p><w:r><w:delText>Moved from row</w:delText></w:r></w:p></w:tc></w:tr>
                      </w:moveFrom>
                      <w:moveTo w:id="44" w:author="D" w:date="2026-06-10T00:00:00Z">
                        <w:tr><w:tc><w:p><w:r><w:t>Moved to row</w:t></w:r></w:p></w:tc></w:tr>
                      </w:moveTo>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteTableMarkupScenarioProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr>
                        <w:trPr><w:tblHeader/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Header two</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
                          <w:p>
                            <w:commentRangeStart w:id="1"/>
                            <w:r><w:t>Commented cell</w:t></w:r>
                            <w:commentRangeEnd w:id="1"/>
                            <w:r><w:commentReference w:id="1"/></w:r>
                          </w:p>
                        </w:tc>
                        <w:tc><w:p><w:r><w:t>Base cell</w:t></w:r></w:p></w:tc>
                        <w:ins w:id="401" w:author="A" w:date="2026-06-10T00:00:00Z">
                          <w:tc><w:p><w:r><w:t>Inserted cell</w:t></w:r></w:p></w:tc>
                        </w:ins>
                        <w:del w:id="402" w:author="B" w:date="2026-06-10T00:00:00Z">
                          <w:tc><w:p><w:r><w:delText>Deleted cell</w:delText></w:r></w:p></w:tc>
                        </w:del>
                      </w:tr>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:vMerge/></w:tcPr>
                          <w:p><w:r><w:t>Continuation cell</w:t></w:r></w:p>
                          <w:del w:id="403" w:author="C" w:date="2026-06-10T00:00:00Z">
                            <w:p><w:r><w:delText>Deleted cell paragraph</w:delText></w:r></w:p>
                          </w:del>
                        </w:tc>
                        <w:moveFrom w:id="404" w:author="D" w:date="2026-06-10T00:00:00Z">
                          <w:tc>
                            <w:tbl>
                              <w:tr><w:tc><w:p><w:r><w:delText>Moved nested old</w:delText></w:r></w:p></w:tc></w:tr>
                            </w:tbl>
                          </w:tc>
                        </w:moveFrom>
                        <w:moveTo w:id="405" w:author="E" w:date="2026-06-10T00:00:00Z">
                          <w:tc>
                            <w:tbl>
                              <w:tr><w:tc><w:p><w:r><w:t>Moved nested new</w:t></w:r></w:p></w:tc></w:tr>
                            </w:tbl>
                          </w:tc>
                        </w:moveTo>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="1" w:author="Reviewer" w:date="2026-06-10T00:00:00Z">
                    <w:p><w:r><w:t>Table comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """
        });
    }

    internal static string WriteBodyElementRevisionProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                    <w:ins w:id="101" w:author="A" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:t>Inserted body paragraph</w:t></w:r></w:p>
                    </w:ins>
                    <w:del w:id="102" w:author="B" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:delText>Deleted body paragraph</w:delText></w:r></w:p>
                    </w:del>
                    <w:ins w:id="103" w:author="C" w:date="2026-06-10T00:00:00Z">
                      <w:tbl>
                        <w:tr><w:tc><w:p><w:r><w:t>Inserted body table</w:t></w:r></w:p></w:tc></w:tr>
                      </w:tbl>
                    </w:ins>
                    <w:del w:id="104" w:author="D" w:date="2026-06-10T00:00:00Z">
                      <w:tbl>
                        <w:tr><w:tc><w:p><w:r><w:delText>Deleted body table</w:delText></w:r></w:p></w:tc></w:tr>
                      </w:tbl>
                    </w:del>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteBlockContentControlFormattingRevisionProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                    <w:sdt>
                      <w:sdtPr><w:alias w:val="Synthetic block control"/></w:sdtPr>
                      <w:sdtContent>
                        <w:p>
                          <w:pPr>
                            <w:pPrChange w:id="201" w:author="A" w:date="2026-06-10T00:00:00Z">
                              <w:pPr><w:pStyle w:val="ControlledStyle"/></w:pPr>
                            </w:pPrChange>
                          </w:pPr>
                          <w:r><w:t>Controlled paragraph</w:t></w:r>
                        </w:p>
                      </w:sdtContent>
                    </w:sdt>
                    <w:p><w:r><w:t>After</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteStaticAndRelatedStoryRevisionProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Body</w:t></w:r>
                      <w:r><w:commentReference w:id="1"/></w:r>
                      <w:r><w:footnoteReference w:id="2"/></w:r>
                      <w:r><w:endnoteReference w:id="3"/></w:r>
                    </w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader"/>
                      <w:footerReference w:type="default" r:id="rIdFooter"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p>
                    <w:r><w:t>Header base </w:t></w:r>
                    <w:ins w:id="11" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>Header inserted</w:t></w:r></w:ins>
                    <w:del w:id="12" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>Header deleted</w:delText></w:r></w:del>
                  </w:p>
                </w:hdr>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p>
                    <w:r><w:t>Footer base </w:t></w:r>
                    <w:ins w:id="21" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>Footer inserted</w:t></w:r></w:ins>
                    <w:del w:id="22" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>Footer deleted</w:delText></w:r></w:del>
                  </w:p>
                </w:ftr>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="1">
                    <w:p>
                      <w:r><w:t>Comment base </w:t></w:r>
                      <w:ins w:id="31" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>Comment inserted</w:t></w:r></w:ins>
                      <w:del w:id="32" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>Comment deleted</w:delText></w:r></w:del>
                    </w:p>
                  </w:comment>
                </w:comments>
                """,
            ["word/footnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:id="2">
                    <w:p>
                      <w:r><w:t>Footnote base </w:t></w:r>
                      <w:ins w:id="41" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>Footnote inserted</w:t></w:r></w:ins>
                      <w:del w:id="42" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>Footnote deleted</w:delText></w:r></w:del>
                    </w:p>
                  </w:footnote>
                </w:footnotes>
                """,
            ["word/endnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:endnote w:id="3">
                    <w:p>
                      <w:r><w:t>Endnote base </w:t></w:r>
                      <w:ins w:id="51" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>Endnote inserted</w:t></w:r></w:ins>
                      <w:del w:id="52" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>Endnote deleted</w:delText></w:r></w:del>
                    </w:p>
                  </w:endnote>
                </w:endnotes>
                """
        });
    }

    internal static string WriteSectionHeaderFooterMarkupProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
                  <Override PartName="/word/headerDefault.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/headerEven.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/headerFirst.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footerDefault.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                  <Override PartName="/word/footerEven.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                  <Override PartName="/word/footerFirst.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
                  <Relationship Id="rIdHeaderDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerDefault.xml"/>
                  <Relationship Id="rIdHeaderEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerEven.xml"/>
                  <Relationship Id="rIdHeaderFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerFirst.xml"/>
                  <Relationship Id="rIdFooterDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerDefault.xml"/>
                  <Relationship Id="rIdFooterEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerEven.xml"/>
                  <Relationship Id="rIdFooterFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerFirst.xml"/>
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                </Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:evenAndOddHeaders/>
                </w:settings>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
                    <w:ins w:id="801" w:author="A" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:br w:type="page"/></w:r></w:p>
                    </w:ins>
                    <w:del w:id="802" w:author="B" w:date="2026-06-10T00:00:00Z">
                      <w:p><w:r><w:br w:type="column"/></w:r></w:p>
                    </w:del>
                    <w:ins w:id="803" w:author="C" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:pPr>
                          <w:sectPr>
                            <w:type w:val="nextPage"/>
                            <w:headerReference w:type="default" r:id="rIdHeaderDefault"/>
                            <w:headerReference w:type="even" r:id="rIdHeaderEven"/>
                            <w:headerReference w:type="first" r:id="rIdHeaderFirst"/>
                            <w:footerReference w:type="default" r:id="rIdFooterDefault"/>
                            <w:footerReference w:type="even" r:id="rIdFooterEven"/>
                            <w:footerReference w:type="first" r:id="rIdFooterFirst"/>
                          </w:sectPr>
                        </w:pPr>
                      </w:p>
                    </w:ins>
                    <w:p><w:r><w:t>After</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeaderDefault"/>
                      <w:headerReference w:type="even" r:id="rIdHeaderEven"/>
                      <w:headerReference w:type="first" r:id="rIdHeaderFirst"/>
                      <w:footerReference w:type="default" r:id="rIdFooterDefault"/>
                      <w:footerReference w:type="even" r:id="rIdFooterEven"/>
                      <w:footerReference w:type="first" r:id="rIdFooterFirst"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/headerDefault.xml"] = StaticStoryPartXml("hdr", "Default header"),
            ["word/headerEven.xml"] = StaticStoryPartXml("hdr", "Even header"),
            ["word/headerFirst.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p>
                    <w:ins w:id="811" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>First header inserted</w:t></w:r></w:ins>
                    <w:del w:id="812" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>First header deleted</w:delText></w:r></w:del>
                  </w:p>
                  <w:p>
                    <w:commentRangeStart w:id="1"/>
                    <w:r><w:t>First header comment anchor</w:t></w:r>
                    <w:commentRangeEnd w:id="1"/>
                    <w:r><w:commentReference w:id="1"/></w:r>
                  </w:p>
                </w:hdr>
                """,
            ["word/footerDefault.xml"] = StaticStoryPartXml("ftr", "Default footer"),
            ["word/footerEven.xml"] = StaticStoryPartXml("ftr", "Even footer"),
            ["word/footerFirst.xml"] = StaticStoryPartXml("ftr", "First footer"),
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="1" w:author="Reviewer" w:date="2026-06-10T00:00:00Z">
                    <w:p><w:r><w:t>Static story comment</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """
        });

        static string StaticStoryPartXml(string root, string label)
        {
            return $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:{{root}} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p>
                    <w:ins w:id="810" w:author="A" w:date="2026-06-10T00:00:00Z"><w:r><w:t>{{label}} inserted</w:t></w:r></w:ins>
                    <w:del w:id="820" w:author="B" w:date="2026-06-10T00:00:00Z"><w:r><w:delText>{{label}} deleted</w:delText></w:r></w:del>
                  </w:p>
                </w:{{root}}>
                """;
        }
    }

    internal static string WriteImageRevisionProbeDocx()
    {
        byte[] png = TestFixtures.CreateRgbPng(1, 1, [16, 32, 48]);
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
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
                  <Relationship Id="rIdInlineInserted" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inline-inserted.png"/>
                  <Relationship Id="rIdInlineDeleted" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inline-deleted.png"/>
                  <Relationship Id="rIdFloatingInserted" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/floating-inserted.png"/>
                  <Relationship Id="rIdFloatingDeleted" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/floating-deleted.png"/>
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
                      <w:ins w:id="61" w:author="A" w:date="2026-06-10T00:00:00Z">
                        <w:r>
                          <w:drawing>
                            <wp:inline>
                              <wp:extent cx="914400" cy="914400"/>
                              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdInlineInserted"/></pic:blipFill></pic:pic>
                              </a:graphicData></a:graphic>
                            </wp:inline>
                          </w:drawing>
                        </w:r>
                      </w:ins>
                      <w:del w:id="62" w:author="B" w:date="2026-06-10T00:00:00Z">
                        <w:r>
                          <w:drawing>
                            <wp:inline>
                              <wp:extent cx="914400" cy="914400"/>
                              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdInlineDeleted"/></pic:blipFill></pic:pic>
                              </a:graphicData></a:graphic>
                            </wp:inline>
                          </w:drawing>
                        </w:r>
                      </w:del>
                      <w:ins w:id="63" w:author="C" w:date="2026-06-10T00:00:00Z">
                        <w:r>
                          <w:drawing>
                            <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="0">
                              <wp:extent cx="914400" cy="914400"/>
                              <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                              <wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>
                              <wp:wrapNone/>
                              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdFloatingInserted"/></pic:blipFill></pic:pic>
                              </a:graphicData></a:graphic>
                            </wp:anchor>
                          </w:drawing>
                        </w:r>
                      </w:ins>
                      <w:del w:id="64" w:author="D" w:date="2026-06-10T00:00:00Z">
                        <w:r>
                          <w:drawing>
                            <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="0">
                              <wp:extent cx="914400" cy="914400"/>
                              <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                              <wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>
                              <wp:wrapNone/>
                              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdFloatingDeleted"/></pic:blipFill></pic:pic>
                              </a:graphicData></a:graphic>
                            </wp:anchor>
                          </w:drawing>
                        </w:r>
                      </w:del>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/inline-inserted.png"] = png,
            ["word/media/inline-deleted.png"] = png,
            ["word/media/floating-inserted.png"] = png,
            ["word/media/floating-deleted.png"] = png
        });
    }

    internal static string WriteMovedCommentedImageProbeDocx()
    {
        byte[] png = TestFixtures.CreateRgbPng(1, 1, [80, 96, 112]);
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rIdBodyCommented" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/body-commented.png"/>
                  <Relationship Id="rIdInlineMoveFrom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inline-move-from.png"/>
                  <Relationship Id="rIdInlineMoveTo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/inline-move-to.png"/>
                  <Relationship Id="rIdFloatingMoveFrom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/floating-move-from.png"/>
                  <Relationship Id="rIdFloatingMoveTo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/floating-move-to.png"/>
                </Relationships>
                """),
            ["word/_rels/comments.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdCommentAnchor" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/comment-anchor.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8($$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:commentRangeStart w:id="1"/>
                      {{InlineDrawingXml("rIdBodyCommented")}}
                      <w:commentRangeEnd w:id="1"/>
                      <w:r><w:commentReference w:id="1"/></w:r>
                      <w:moveFrom w:id="701" w:author="A" w:date="2026-06-10T00:00:00Z">
                        {{InlineDrawingXml("rIdInlineMoveFrom")}}
                      </w:moveFrom>
                      <w:moveTo w:id="702" w:author="B" w:date="2026-06-10T00:00:00Z">
                        {{InlineDrawingXml("rIdInlineMoveTo")}}
                      </w:moveTo>
                      <w:moveFrom w:id="703" w:author="C" w:date="2026-06-10T00:00:00Z">
                        {{AnchorDrawingXml("rIdFloatingMoveFrom", "0")}}
                      </w:moveFrom>
                      <w:moveTo w:id="704" w:author="D" w:date="2026-06-10T00:00:00Z">
                        {{AnchorDrawingXml("rIdFloatingMoveTo", "1")}}
                      </w:moveTo>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/comments.xml"] = TestFixtures.Utf8($$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:comment w:id="1" w:author="Reviewer" w:date="2026-06-10T00:00:00Z">
                    <w:p>
                      <w:moveTo w:id="705" w:author="E" w:date="2026-06-10T00:00:00Z">
                        {{AnchorDrawingXml("rIdCommentAnchor", "1")}}
                      </w:moveTo>
                    </w:p>
                  </w:comment>
                </w:comments>
                """),
            ["word/media/body-commented.png"] = png,
            ["word/media/inline-move-from.png"] = png,
            ["word/media/inline-move-to.png"] = png,
            ["word/media/floating-move-from.png"] = png,
            ["word/media/floating-move-to.png"] = png,
            ["word/media/comment-anchor.png"] = png
        });

        static string InlineDrawingXml(string relationshipId)
        {
            return $$"""
                <w:r>
                  <w:drawing>
                    <wp:inline>
                      <wp:extent cx="914400" cy="914400"/>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic><pic:blipFill><a:blip r:embed="{{relationshipId}}"/></pic:blipFill></pic:pic>
                      </a:graphicData></a:graphic>
                    </wp:inline>
                  </w:drawing>
                </w:r>
                """;
        }

        static string AnchorDrawingXml(string relationshipId, string behindDocumentValue)
        {
            return $$"""
                <w:r>
                  <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="{{behindDocumentValue}}">
                      <wp:extent cx="914400" cy="914400"/>
                      <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                      <wp:positionV relativeFrom="paragraph"><wp:posOffset>914400</wp:posOffset></wp:positionV>
                      <wp:wrapNone/>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic><pic:blipFill><a:blip r:embed="{{relationshipId}}"/></pic:blipFill></pic:pic>
                      </a:graphicData></a:graphic>
                    </wp:anchor>
                  </w:drawing>
                </w:r>
                """;
        }
    }

    internal static string WriteCommentMarkerProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                      <w:r><w:t>Commented</w:t></w:r>
                      <w:r><w:commentReference w:id="1"/></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteCommentAnchorAccountingProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Visible anchor</w:t></w:r>
                      <w:r><w:commentReference w:id="1"/></w:r>
                    </w:p>
                    <w:del w:id="10" w:author="A" w:date="2026-06-10T00:00:00Z">
                      <w:p>
                        <w:r><w:delText>Hidden anchor</w:delText></w:r>
                        <w:r><w:commentReference w:id="2"/></w:r>
                      </w:p>
                    </w:del>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="1"><w:p><w:r><w:t>Visible body</w:t></w:r></w:p></w:comment>
                  <w:comment w:id="2"><w:p><w:r><w:t>Hidden body</w:t></w:r></w:p></w:comment>
                  <w:comment w:id="3"><w:p><w:r><w:t>Orphan body</w:t></w:r></w:p></w:comment>
                  <w:comment><w:p><w:r><w:t>Unsupported body</w:t></w:r></w:p></w:comment>
                </w:comments>
                """
        });
    }

    internal static string WriteFormattingRevisionProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                        <w:pPrChange w:id="10" w:author="A" w:date="2026-06-05T00:00:00Z">
                          <w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr>
                        </w:pPrChange>
                        <w:sectPr>
                          <w:pgSz w:w="12240" w:h="15840"/>
                          <w:sectPrChange w:id="12" w:author="C" w:date="2026-06-07T00:00:00Z">
                            <w:sectPr><w:pgMar w:left="1440" w:right="1440"/><w:cols w:num="2"/></w:sectPr>
                          </w:sectPrChange>
                        </w:sectPr>
                      </w:pPr>
                      <w:r>
                        <w:rPr>
                          <w:rPrChange w:id="11" w:author="B" w:date="2026-06-06T00:00:00Z">
                            <w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>
                          </w:rPrChange>
                        </w:rPr>
                        <w:t>Formatting revision</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static string WriteTableFormattingRevisionProbeDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                        <w:tblPrChange w:id="20" w:author="A" w:date="2026-06-07T00:00:00Z">
                          <w:tblPr>
                            <w:tblW w:w="2400" w:type="dxa"/>
                            <w:tblBorders><w:top w:val="single"/></w:tblBorders>
                          </w:tblPr>
                        </w:tblPrChange>
                      </w:tblPr>
                      <w:tr>
                        <w:trPr>
                          <w:trHeight w:val="360"/>
                          <w:trPrChange w:id="21" w:author="B" w:date="2026-06-08T00:00:00Z">
                            <w:trPr><w:trHeight w:val="720"/><w:cantSplit/></w:trPr>
                          </w:trPrChange>
                        </w:trPr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcW w:w="2400" w:type="dxa"/>
                            <w:tcPrChange w:id="22" w:author="C" w:date="2026-06-09T00:00:00Z">
                              <w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:shd w:fill="FFFF00"/></w:tcPr>
                            </w:tcPrChange>
                          </w:tcPr>
                          <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    internal static DocxDocument CreateFontPlanDocument(DocxTextRun run, DocxFontCatalog fontCatalog)
    {
        return CreateFontPlanDocument([run], fontCatalog);
    }

    internal static DocxDocument CreateFontPlanDocument(IReadOnlyList<DocxTextRun> runs, DocxFontCatalog fontCatalog)
    {
        return CreateFontPlanDocument([CreateFontPlanParagraph(runs)], fontCatalog);
    }

    internal static DocxParagraph CreateFontPlanParagraph(DocxTextRun run)
    {
        return CreateFontPlanParagraph([run]);
    }

    internal static DocxParagraph CreateFontPlanParagraph(IReadOnlyList<DocxTextRun> runs)
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

    internal static DocxDocument CreateFontPlanDocument(IReadOnlyList<DocxParagraph> paragraphs, DocxFontCatalog fontCatalog)
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

    internal static (FontFaceResolution Resolution, string FirstFamily)? TryLoadCollectionFace(FontFaceResolution resolution)
    {
        try
        {
            ReadOnlyMemory<byte> bytes = resolution.Source.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult();
            OpenTypeFont first = OpenTypeFont.Load(bytes.ToArray(), 0);
            OpenTypeFont selected = OpenTypeFont.Load(bytes.ToArray(), resolution.FontFaceIndex);
            return selected.FamilyName.Equals(resolution.FamilyName, StringComparison.OrdinalIgnoreCase)
                ? (resolution, first.FamilyName)
                : null;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    internal static (FontFaceResolution Resolution, OpenTypeFont Font)? FindUsableInstalledFont()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return null;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        foreach (FontFaceResolution resolution in resolver.GetDiscoveredFonts())
        {
            try
            {
                ReadOnlyMemory<byte> bytes = resolution.Source.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult();
                OpenTypeFont font = OpenTypeFont.Load(bytes.ToArray(), resolution.FontFaceIndex);
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

    internal static (FontFaceResolution Resolution, OpenTypeFont Font)? FindUsableInstalledFontExcept(string familyName)
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            return null;
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        foreach (FontFaceResolution resolution in resolver.GetDiscoveredFonts().Where(font => !font.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase)))
        {
            try
            {
                ReadOnlyMemory<byte> bytes = resolution.Source.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult();
                OpenTypeFont font = OpenTypeFont.Load(bytes.ToArray(), resolution.FontFaceIndex);
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

    internal static double MeasureOpenTypeText(OpenTypeFont font, string text, double fontSize)
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

    internal sealed class SingleResolutionFontResolver(FontFaceResolution resolution) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return resolution with
            {
                Style = resolution.Style with { Bold = request.Bold, Italic = request.Italic },
                RequestedFamily = request.FamilyName,
                IsFallback = !request.FamilyName.Equals(resolution.FamilyName, StringComparison.OrdinalIgnoreCase)
            };
        }
    }

    internal sealed class CffPrimaryFontResolver(byte[] cffBytes, byte[] fallbackBytes, bool emitCff = true) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            bool cff = emitCff && request.FamilyName.Equals("CffFamily", StringComparison.OrdinalIgnoreCase);
            var source = new MemoryFontProgramSource(cff ? "cff-test" : "tt-test", cff ? cffBytes : fallbackBytes);
            return new FontFaceResolution(
                request.FamilyName,
                cff ? "CffFamily" : "FallbackFamily",
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                source,
                IsFallback: false);
        }
    }


    internal sealed class FamilyWidthTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
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

    // Mirrors the production scaled measurer (advances times the layout scale, line metrics
    // times the fitted compromise) so Word-compatible layout tests model production instead of
    // pairing raw advances with scaled table geometry (W6-a1).
    internal sealed class ScaledLayoutTextMeasurer(IDocxTextMeasurer inner, double textScale, double lineMetricScale) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            return inner.MeasureText(run, text, fontSize) * textScale;
        }

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxLineMetricsProvider lineMetrics
                ? lineMetrics.MeasureSingleLineHeight(run, fontSize) * lineMetricScale
                : fontSize * 1.2d * lineMetricScale;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsAscender(run, fontSize) * lineMetricScale
                : fontSize * lineMetricScale;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsDescender(run, fontSize) * lineMetricScale
                : fontSize * 0.2d * lineMetricScale;
        }
    }

    internal sealed class FractionalLineHeightTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
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

    internal sealed class FontSizeWidthTextMeasurer : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
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

    internal static DocxTable CreateSingleCellTable(string text, double rowHeight)
    {
        return new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell(text, [], null, null, null, null, [], DocxTableCellMargins.Empty)], rowHeight)]);
    }

    internal static DocxDocument CreateLayoutTestDocument(IReadOnlyList<DocxBodyElement> bodyElements, IReadOnlyList<DocxTable> tables)
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

    internal static DocxDocument CreateAllMarkupWrapProbeDocument(IReadOnlyList<DocxParagraph> paragraphs)
    {
        return new DocxDocument(
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
            paragraphs.Select(paragraph => new DocxParagraphElement(paragraph)).ToArray(),
            paragraphs,
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
    }

    internal static DocxParagraph CreateDocxLayoutParagraph(
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

    internal static double[] ExtractTextBaselines(string pdf)
    {
        return Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .Distinct()
            .ToArray();
    }

    internal static double[] ExtractTextBaselinesAtX(string pdf, double x)
    {
        string escapedX = Regex.Escape(x.ToString("0.###", CultureInfo.InvariantCulture));
        return Regex.Matches(pdf, $@"1 0 0 1 {escapedX}(?:\.\d+)? (?<y>-?\d+(?:\.\d+)?) Tm")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .Distinct()
            .ToArray();
    }

    internal static int CountPdfTextShows(string pdf)
    {
        return CountOccurrences(pdf, "> Tj") + CountOccurrences(pdf, "] TJ");
    }

    internal static bool HasCommentReferenceSpacerTextOperation(DocxTextEmissionSnapshot snapshot)
    {
        return snapshot.Lines
            .Where(line => line.CommentReferenceCount != 0)
            .SelectMany(line => line.Segments)
            .Any(segment =>
                !segment.IsTerminalLineSpace &&
                segment.Width > 0d &&
                segment.TextLength != 0 &&
                segment.CharacterProfile.WhitespaceCount == segment.TextLength);
    }

    internal static double FirstTextEmissionX(DocxTextEmissionSnapshot snapshot)
    {
        return snapshot.Lines
            .SelectMany(line => line.Segments)
            .Where(segment => !segment.IsTerminalLineSpace)
            .Select(segment => segment.X)
            .First();
    }

    internal static int FirstPdfTextShowIndex(string pdf)
    {
        int tj = pdf.IndexOf("> Tj", StringComparison.Ordinal);
        int positioned = pdf.IndexOf("] TJ", StringComparison.Ordinal);
        return MinNonNegative(tj, positioned);
    }

    internal static int LastPdfTextShowIndex(string pdf)
    {
        return Math.Max(
            pdf.LastIndexOf("> Tj", StringComparison.Ordinal),
            pdf.LastIndexOf("] TJ", StringComparison.Ordinal));
    }

    internal static int CountOccurrences(string text, string value)
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
