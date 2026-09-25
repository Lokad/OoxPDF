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

internal static class DocxPageMarkupTests
{
    public static void DocxStaticHeaderRendersMixedRunColorsSeparately()
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
        TestAssert.Equal(5, DocxTests.CountPdfTextShows(pdf));
    }

    public static void DocxSyntheticFooterPageFieldsUseGeneratedPageNumbers()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
        TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 51, "Footer PAGE and NUMPAGES fields should render on each generated page.");
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

    public static void DocxMarkupModesFilterRevisedParagraphKeepRulesAndPageBreakBefore()
    {
        string input = DocxTests.WriteBlockRevisionParagraphPropertiesProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument simpleDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.SimpleMarkup);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("Before|Inserted keep|Moved to keep|After", DocxTests.ParagraphTexts(finalDocument));
        TestAssert.Equal("Before|Deleted keep|Moved from keep|After", DocxTests.ParagraphTexts(originalDocument));
        TestAssert.Equal("Before|Inserted keep|Moved to keep|After", DocxTests.ParagraphTexts(simpleDocument));
        TestAssert.Equal("Before|Inserted keep|Deleted keep|Moved from keep|Moved to keep|After", DocxTests.ParagraphTexts(allDocument));

        AssertKeepRules(finalDocument, "Inserted keep", DocxRevisionKind.Insertion, "611", expectedWidowControl: false);
        AssertKeepRules(finalDocument, "Moved to keep", DocxRevisionKind.MoveTo, "614", expectedWidowControl: true);
        AssertKeepRules(originalDocument, "Deleted keep", DocxRevisionKind.Deletion, "612", expectedWidowControl: false);
        AssertKeepRules(originalDocument, "Moved from keep", DocxRevisionKind.MoveFrom, "613", expectedWidowControl: true);
        AssertKeepRules(allDocument, "Inserted keep", DocxRevisionKind.Insertion, "611", expectedWidowControl: false);
        AssertKeepRules(allDocument, "Deleted keep", DocxRevisionKind.Deletion, "612", expectedWidowControl: false);
        AssertKeepRules(allDocument, "Moved from keep", DocxRevisionKind.MoveFrom, "613", expectedWidowControl: true);
        AssertKeepRules(allDocument, "Moved to keep", DocxRevisionKind.MoveTo, "614", expectedWidowControl: true);

        AssertPageBreakRevisions(finalDocument, ["Insertion:611"]);
        AssertPageBreakRevisions(originalDocument, ["Deletion:612"]);
        AssertPageBreakRevisions(simpleDocument, ["Insertion:611"]);
        AssertPageBreakRevisions(allDocument, ["Insertion:611", "Deletion:612"]);

        TestAssert.Equal(2, new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(finalDocument).Pages.Count);
        TestAssert.Equal(2, new DocxRenderer(null, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(originalDocument).Pages.Count);
        TestAssert.Equal(2, new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(simpleDocument).Pages.Count);
        TestAssert.Equal(3, new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(allDocument).Pages.Count);

        static void AssertKeepRules(
            DocxDocument document,
            string text,
            DocxRevisionKind expectedRevisionKind,
            string expectedRevisionId,
            bool expectedWidowControl)
        {
            DocxParagraph paragraph = document.Paragraphs.Single(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)) == text);
            TestAssert.True(
                paragraph.Revisions.Any(revision => revision.Kind == expectedRevisionKind && revision.Id == expectedRevisionId),
                "Filtered revised paragraphs should retain block-level revision provenance.");
            TestAssert.True(
                paragraph.KeepRules.KeepNext == true &&
                paragraph.KeepRules.KeepLines == true &&
                paragraph.KeepRules.WidowControl == expectedWidowControl,
                "Filtered revised paragraphs should retain their paragraph keep and widow/orphan rules.");
        }

        static void AssertPageBreakRevisions(DocxDocument document, IReadOnlyList<string> expected)
        {
            string actual = string.Join(
                "|",
                document.BodyElements
                    .OfType<DocxPageBreakElement>()
                    .Select(element => $"{element.Revisions.Single().Kind}:{element.Revisions.Single().Id}"));
            TestAssert.Equal(string.Join("|", expected), actual);
        }
    }

    public static void DocxMarkupModesFilterSectionHeaderFooterAndBreakRevisions()
    {
        string input = DocxTests.WriteSectionHeaderFooterMarkupProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        string finalStaticText = StaticStoryText(finalDocument);
        string originalStaticText = StaticStoryText(originalDocument);
        string allStaticText = StaticStoryText(allDocument);

        AssertMarkupStoryText(
            finalStaticText,
            visible: ["Default header inserted", "Even header inserted", "First header inserted", "Default footer inserted", "Even footer inserted", "First footer inserted"],
            hidden: ["Default header deleted", "Even header deleted", "First header deleted", "Default footer deleted", "Even footer deleted", "First footer deleted"]);
        AssertMarkupStoryText(
            originalStaticText,
            visible: ["Default header deleted", "Even header deleted", "First header deleted", "Default footer deleted", "Even footer deleted", "First footer deleted"],
            hidden: ["Default header inserted", "Even header inserted", "First header inserted", "Default footer inserted", "Even footer inserted", "First footer inserted"]);
        AssertMarkupStoryText(
            allStaticText,
            visible: ["Default header inserted", "Default header deleted", "Even header inserted", "Even header deleted", "First header inserted", "First header deleted", "Default footer inserted", "Default footer deleted", "Even footer inserted", "Even footer deleted", "First footer inserted", "First footer deleted"],
            hidden: []);

        TestAssert.True(StaticStoryParagraphs(allDocument).Any(paragraph => paragraph.CommentRanges.Any(range => range.Id == "1") && paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment && reference.Id == "1")), "Comments anchored in first/even/default static stories should survive all-markup filtering.");
        TestAssert.True(finalDocument.BodyElements.OfType<DocxPageBreakElement>().Any(element => element.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Insertion && revision.Id == "801")), "Inserted run page breaks should survive final view as revision-provenance break elements.");
        TestAssert.True(!finalDocument.BodyElements.OfType<DocxManualBreakElement>().Any(element => element.Value == "column" && element.Revisions.Any(revision => revision.Id == "802")), "Deleted column breaks should be hidden from final view.");
        TestAssert.True(originalDocument.BodyElements.OfType<DocxManualBreakElement>().Any(element => element.Value == "column" && element.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion && revision.Id == "802")), "Deleted run column breaks should survive original view as revision-provenance break elements.");
        TestAssert.True(!originalDocument.BodyElements.OfType<DocxSectionBreakElement>().Any(element => element.Revisions.Any(revision => revision.Id == "803")), "Inserted section breaks should be hidden from original view.");
        TestAssert.True(allDocument.BodyElements.OfType<DocxSectionBreakElement>().Any(element => element.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Insertion && revision.Id == "803")), "Inserted section breaks should retain inherited revision provenance in all-markup view.");

        static string StaticStoryText(DocxDocument document)
        {
            return string.Join("|", StaticStoryParagraphs(document).Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
        }

        static IEnumerable<DocxParagraph> StaticStoryParagraphs(DocxDocument document)
        {
            return DocxBlockTraversal
                .EnumerateStaticStoryParagraphs(document.HeaderBodyElementsByType, document.HeaderParagraphsByType)
                .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.FooterBodyElementsByType, document.FooterParagraphsByType));
        }

        static void AssertMarkupStoryText(string text, IReadOnlyList<string> visible, IReadOnlyList<string> hidden)
        {
            foreach (string expected in visible)
            {
                TestAssert.Contains(expected, text);
            }

            foreach (string unexpected in hidden)
            {
                TestAssert.DoesNotContain(unexpected, text);
            }
        }
    }

    public static void DocxMarkupReserveMarginUsesSectionSpecificPageGeometry()
    {
        DocxPageSettings firstSectionSettings = new(
            "8000",
            "12000",
            null,
            "720",
            "720",
            "720",
            "720",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxPageSettings finalSectionSettings = new(
            "12240",
            "15840",
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
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            finalSectionSettings,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Narrow first section", 10d, 12d)),
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide final section", 10d, 12d))
            ],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxLayoutSnapshot layout = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .InspectLayout(document);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.True(Math.Abs(layout.Pages[0].Width - 400d) < 0.001d, "The first section should own its authored narrow page width.");
        TestAssert.True(Math.Abs(layout.Pages[0].MarginRight - 148d) < 0.001d, "The narrow first section should cap the reserved markup margin at the minimum body width.");
        TestAssert.True(Math.Abs(layout.Pages[0].MarkupMarginReservePoints - 112d) < 0.001d, "The narrow first section should report only the reserve added beyond its own authored margin.");
        TestAssert.True(Math.Abs(layout.Pages[1].Width - 612d) < 0.001d, "The final section should own its authored Letter page width.");
        TestAssert.True(Math.Abs(layout.Pages[1].MarginRight - 207d) < 0.001d, "The wider final section should reach the preferred review margin target.");
        TestAssert.True(Math.Abs(layout.Pages[1].MarkupMarginReservePoints - 135d) < 0.001d, "The final section should report the reserve added beyond its own authored margin.");
    }

    public static void DocxMarkupReserveMarginKeepsRightLaneOnMirroredEvenPages()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "docx-markup-margin-mirrored.docx"));
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var preserveRenderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var reserveRenderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin);

        DocxLayoutSnapshot preserveLayout = preserveRenderer.InspectLayout(allDocument);
        DocxLayoutSnapshot reserveLayout = reserveRenderer.InspectLayout(allDocument);
        DocxMarkupBalloonPlacementSnapshot[] placements = reserveRenderer.InspectMarkupBalloons(allDocument).ToArray();

        TestAssert.True(allDocument.Settings.MirrorMargins == true, "The mirrored-margin fixture should expose the modeled document setting.");
        TestAssert.True(reserveLayout.Pages.Count >= 2, "The mirrored-margin fixture should cover odd and even pages.");
        TestAssert.True(reserveLayout.Pages[0].MarginRight > preserveLayout.Pages[0].MarginRight, "Odd mirrored pages should keep the reserved markup lane on the right.");
        TestAssert.Equal(preserveLayout.Pages[0].MarginLeft, reserveLayout.Pages[0].MarginLeft);
        TestAssert.True(reserveLayout.Pages[1].MarginRight > preserveLayout.Pages[1].MarginRight, "Even mirrored pages should keep the reserved markup lane on the right (Office balloons right on both pages).");
        TestAssert.Equal(preserveLayout.Pages[1].MarginLeft, reserveLayout.Pages[1].MarginLeft);
        TestAssert.True(placements.Any(placement => placement.PageIndex == 0 && placement.Side == "Right"), "Odd-page balloons should use the right mirrored lane.");
        TestAssert.True(placements.Any(placement => placement.PageIndex == 1 && placement.Side == "Right"), "Even-page balloons should use the right lane like odd pages.");
        TestAssert.True(
            placements.Where(placement => placement.Side == "Right").All(placement => placement.BalloonConnectorX < placement.X),
            "Right-lane connector stems should stay between the body frame and balloon body.");
    }

    public static void DocxWordCompatibleAllMarkupPaintsPageRevisionBar()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Revision bar public probe", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "Reviewer", "2026-06-01T00:00:00Z", "inserted", null, [])
            ]
        };
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
            [new DocxParagraphElement(paragraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0 g", page.Content);
        TestAssert.Contains("27.925 ", page.Content);
        TestAssert.Contains(" 0.475 ", page.Content);
        TestAssert.True(
            !page.Content.Contains(" 1.5 ", StringComparison.Ordinal),
            "Word-compatible all-markup should use a single page revision bar instead of legacy per-line change bars.");
    }

    public static void DocxWordCompatibleAllMarkupPaintsPageRevisionBarForRunRevisions()
    {
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "11", "Reviewer", "2026-06-01T00:00:00Z", "inserted", null, []);
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Revision bar run probe", 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun("Revision ", 10d, null, false, false, false, null, null),
                new DocxTextRun("run", 10d, null, false, false, false, null, null)
                {
                    Revision = revision
                },
                new DocxTextRun(" probe", 10d, null, false, false, false, null, null)
            ]
        };
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
            [new DocxParagraphElement(paragraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0 g", page.Content);
        TestAssert.Contains("27.925 ", page.Content);
        TestAssert.Contains(" 0.475 ", page.Content);
    }

    public static void DocxWordCompatibleAllMarkupPaintsPageRevisionBarForPlacedFootnoteRevisions()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "27",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote revised text", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "27", "Reviewer", "2026-06-10T00:00:00Z", "inserted", null, [])
            ]
        };
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "27",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
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
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            RelatedStories = [footnoteStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0 g", page.Content);
        TestAssert.Contains("27.925 ", page.Content);
        TestAssert.Contains(" 0.475 ", page.Content);
        TestAssert.True(
            !page.Content.Contains(" 1.5 ", StringComparison.Ordinal),
            "Word-compatible all-markup should keep placed footnote revisions on the shared page revision bar path.");
    }

    public static void DocxWordCompatibleAllMarkupPaintsPageRevisionBarForPlacedEndnoteRevisions()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "28",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote revised text", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "28", "Reviewer", "2026-06-10T00:00:00Z", "inserted", null, [])
            ]
        };
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "28",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
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
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            RelatedStories = [endnoteStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0 g", page.Content);
        TestAssert.Contains("27.925 ", page.Content);
        TestAssert.Contains(" 0.475 ", page.Content);
        TestAssert.True(
            !page.Content.Contains(" 1.5 ", StringComparison.Ordinal),
            "Word-compatible all-markup should keep placed endnote revisions on the shared page revision bar path.");
    }

    public static void DocxWordCompatibleAllMarkupPaintsPageRevisionBarForFloatingTextBoxRunRevisions()
    {
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "37", "Reviewer", "2026-06-10T00:00:00Z", "inserted", null, []);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph("Text box revised run", 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun("Text box ", 10d, null, false, false, false, null, null),
                new DocxTextRun("revised", 10d, null, false, false, false, null, null)
                {
                    Revision = revision
                },
                new DocxTextRun(" run", 10d, null, false, false, false, null, null)
            ]
        };
        DocxFloatingDrawing drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [drawing],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d))],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0 g", page.Content);
        TestAssert.Contains("27.925 ", page.Content);
        TestAssert.Contains(" 0.475 ", page.Content);
    }

    public static void DocxWordCompatibleAllMarkupPaintsPageRevisionBarForStaticFloatingTextBoxRunRevisions()
    {
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "38", "Reviewer", "2026-06-10T00:00:00Z", "inserted", null, []);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph("Static text box revised run", 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun("Static text box ", 10d, null, false, false, false, null, null),
                new DocxTextRun("revised", 10d, null, false, false, false, null, null)
                {
                    Revision = revision
                },
                new DocxTextRun(" run", 10d, null, false, false, false, null, null)
            ]
        };
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)])]
            }
        };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0 g", page.Content);
        TestAssert.Contains("27.925 ", page.Content);
        TestAssert.Contains(" 0.475 ", page.Content);
    }
}
