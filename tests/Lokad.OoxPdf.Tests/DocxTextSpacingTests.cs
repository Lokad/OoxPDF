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

internal static class DocxTextSpacingTests
{
    public static void DocxParagraphLayoutPreservesAuthoredSpaces()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), " Alpha Beta".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        // RV06 row-end matrix: Office keeps one row-end space beyond authored trailing.
        TestAssert.Equal(" Alpha  Beta  ", line.Text);
        TestAssert.True(line.Width > embedded.MeasureTextPoints("Alpha Beta", 11d), "Preserved spaces should contribute to layout width.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxBodyElement[] elements = document.BodyElements.ToArray();

        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("First", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxManualBreakElement, "A run-level column-break-only paragraph should become a body manual break.");
        var manualBreak = (DocxManualBreakElement)elements[1];
        TestAssert.Equal(DocxBreakSourceKind.RunBreak, manualBreak.SourceKind);
        TestAssert.Equal("column", manualBreak.Value ?? string.Empty);
        TestAssert.True(manualBreak.BreakParagraph is not null, "The authored break paragraph should remain available for future column-flow layout.");
        TestAssert.Equal("Second", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(1, snapshot.ManualBreakBlockCount);
        DocxStructureBlockSnapshot block = snapshot.Blocks[1];
        TestAssert.Equal("ManualBreak", block.Kind);
        TestAssert.Equal("Paragraph", block.PreviousKind ?? string.Empty);
        TestAssert.Equal("Paragraph", block.NextKind ?? string.Empty);
        TestAssert.Equal("runBreak", block.ManualBreakSourceKind ?? string.Empty);
        TestAssert.Equal("column", block.ManualBreakValue ?? string.Empty);
        TestAssert.True(block.ManualBreakConsumesParagraphLine == true, "The snapshot should expose the preserved break paragraph.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxBodyElement[] elements = document.BodyElements.ToArray();
        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("Left", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxManualBreakElement, "The inline column break should become a body manual break.");
        TestAssert.Equal("Right", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(2, snapshot.Pages[0].ColumnFrameCount);
        DocxLayoutItemSnapshot[] lines = snapshot.Pages[0].Items.Where(item => item.Kind == "TextLine").ToArray();
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(0, lines[0].ColumnIndex ?? -1);
        TestAssert.Equal(1, lines[1].ColumnIndex ?? -1);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxSectionBreakElement sectionBreak = document.BodyElements.OfType<DocxSectionBreakElement>().Single();

        TestAssert.Equal("continuous", sectionBreak.TypeValue?.ToValueString() ?? string.Empty);
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

    public static void DocxSyntheticExactLineHeightPositionsNextParagraph()
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
        double[] baselines = DocxTests.ExtractTextBaselines(pdf);
        TestAssert.True(baselines.Length >= 2, "Expected at least two rendered text baselines.");
        TestAssert.True(Math.Abs((baselines[0] - baselines[1]) - 36d) < 0.01d, "Exact DOCX line height should advance the next paragraph by 36 points.");
    }

    public static void DocxLayoutStageAppliesParagraphIndentsToWrapping()
    {
        DocxParagraph indented = DocxTests.CreateDocxLayoutParagraph(
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
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
        var resolver = new DocxTests.SingleResolutionFontResolver(font.Value.Resolution);
        DocxFontPlan plan = DocxFontPlan.Create(document, resolver, CancellationToken.None);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxFontPlanTextMeasurer(plan, null, CancellationToken.None), CancellationToken.None)
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
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("Alpha", 10d, 10d);
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
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Beta", 10d, 10d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(first), new DocxParagraphElement(empty), new DocxParagraphElement(second)],
            []);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FirstSecond".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
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

    public static void DocxSyntheticContextualSpacingSuppressesDefaultStyleGap()
    {
        var spacing = new DocxParagraphSpacing(null, null, null, null, null, null, null, null, true);
        var first = new DocxParagraph(
            [new DocxTextRun("First", 10d, null, false, false, false, null, null)],
            [],
            null,
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(first), new DocxParagraphElement(second)],
            []);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(10d, Math.Round(lines[0].BaselineY - lines[1].BaselineY, 3));
        TestAssert.Equal(10d, lines[1].PendingAfterSpacing ?? -1d);
        TestAssert.Equal(10d, lines[1].ParagraphBeforeSpacing ?? -1d);
        TestAssert.Equal(0d, lines[1].AppliedBeforeSpacing ?? -1d);
        TestAssert.True(lines[1].ContextualSpacingSuppressed == true, "Default-style paragraphs should be treated as same-style paragraphs for contextual spacing.");
    }

    public static void DocxSyntheticContextualSpacingKeepsDifferentStyleGap()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FirstSecond".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
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

    public static void DocxSyntheticAtLeastLineSpacingUsesMinimumAndNaturalLineHeight()
    {
        var smallMinimum = new DocxParagraph(
            [new DocxTextRun("Small", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            6d,
            new DocxParagraphSpacing(null, null, null, null, null, null, "120", "atLeast", null),
            DocxParagraphKeepRules.Empty,
            null);
        var largeMinimum = new DocxParagraph(
            [new DocxTextRun("Large", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            16d,
            new DocxParagraphSpacing(null, null, null, null, null, null, "320", "atLeast", null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(smallMinimum), new DocxParagraphElement(largeMinimum)],
            []);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(10d, lines[0].LineHeight);
        TestAssert.Equal("AtLeastLineSpacing", lines[0].LineHeightSource?.ToString() ?? string.Empty);
        TestAssert.Equal(16d, lines[1].LineHeight);
        TestAssert.Equal("AtLeastLineSpacing", lines[1].LineHeightSource?.ToString() ?? string.Empty);
    }

    public static void DocxSyntheticParagraphKeepNextEstimateHonorsContextualSpacing()
    {
        var contextualSpacing = new DocxParagraphSpacing(null, null, null, null, null, null, null, null, true);
        DocxParagraph CreateBodyParagraph(string text, double before, double after, bool keepNext)
        {
            return new DocxParagraph(
                [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
                [],
                "Body",
                DocxTextAlignment.Left,
                null,
                before,
                after,
                1d,
                10d,
                contextualSpacing,
                keepNext ? new DocxParagraphKeepRules(true, null, null, null, null, null) : DocxParagraphKeepRules.Empty,
                null);
        }

        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
        ];
        DocxParagraph keepNext = CreateBodyParagraph("Keep", before: 0d, after: 12d, keepNext: true);
        DocxParagraph next = CreateBodyParagraph("Next", before: 12d, after: 0d, keepNext: false);
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Concat([new DocxParagraphElement(keepNext), new DocxParagraphElement(next)])
            .ToArray();
        var document = new DocxDocument(
            160d,
            81d,
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages.Single().Items.OfType<DocxTextLineLayout>().ToArray();

        TestAssert.Equal(6, lines.Length);
        TestAssert.Equal("Keep", lines[4].Text);
        TestAssert.Equal("Next", lines[5].Text);
        TestAssert.True(lines[5].ContextualSpacingSuppressed == true, "The kept pair should use the same contextual spacing collapse as normal paragraph layout.");
    }

    public static void DocxSyntheticParagraphExplicitWidowControlOffAllowsWidowLine()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        DocxParagraph widowOff = DocxTests.CreateDocxLayoutParagraph(
            "One\nTwo\nThree",
            fontSize: 10d,
            lineSpacingPoints: 10d,
            keepRules: new DocxParagraphKeepRules(null, null, null, null, false, null));
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FillOneTwoThree".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded, CancellationToken.None);

        TestAssert.Equal(6, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Count());
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
        var defaultAutoSpacing = new DocxParagraphSpacing(null, null, null, null, null, null, null, "auto", null);
        var explicit115Spacing = new DocxParagraphSpacing(null, null, null, null, null, null, "276", "auto", null);
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
            defaultAutoSpacing,
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
            defaultAutoSpacing,
            DocxParagraphKeepRules.Empty,
            label);
        var explicit115List = new DocxParagraph(
            [new DocxTextRun("Explicit 115", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.15d,
            null,
            explicit115Spacing,
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
            explicit115Spacing,
            DocxParagraphKeepRules.Empty,
            null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(flooredList), new DocxParagraphElement(listWithoutBeforeSpacing), new DocxParagraphElement(explicit115List), new DocxParagraphElement(defaultAutoList), new DocxParagraphElement(plainParagraph)],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutItemSnapshot[] textLines = snapshot.Pages[0].Items
            .Where(item => item.Kind == "TextLine")
            .ToArray();

        TestAssert.Equal(5, textLines.Length);
        TestAssert.Equal(10d, textLines[0].SingleLineHeightPoints ?? 0d);
        TestAssert.Equal(0d, textLines[0].PendingAfterSpacingPoints ?? -1d);
        TestAssert.Equal(6d, textLines[0].ParagraphBeforeSpacingPoints ?? 0d);
        TestAssert.Equal(0d, textLines[0].ParagraphAfterSpacingPoints ?? -1d);
        TestAssert.True(textLines[0].ContextualSpacingSuppressed == false, "First paragraph should report that contextual spacing suppression did not apply.");
        TestAssert.Equal(1.16d, textLines[0].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(Math.Abs((textLines[0].LineHeightPoints ?? 0d) - 11.6d) < 0.0001d, "Effective line height should be the measured single-line height multiplied by the effective factor.");
        TestAssert.True(textLines[0].LineSpacingFactorFloorApplied == true, "Default-auto list paragraphs should report the Word-compatible auto-line floor.");
        TestAssert.Equal(0d, textLines[1].PendingAfterSpacingPoints ?? -1d);
        TestAssert.Equal(0d, textLines[1].ParagraphBeforeSpacingPoints ?? -1d);
        TestAssert.Equal(1.16d, textLines[1].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(textLines[1].LineSpacingFactorFloorApplied == true, "List paragraphs report the floor without requiring positive before spacing (Office line-height probe 2026-09-06).");
        TestAssert.Equal(1.15d, textLines[2].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(Math.Abs((textLines[2].LineHeightPoints ?? 0d) - 11.5d) < 0.0001d, "Explicit w:line factors are honored as-authored for lists (Office explicit-115 probe 2026-09-06: explicit 1.15 pitches 14.04 at 10pt Calibri).");
        TestAssert.True(textLines[2].LineSpacingFactorFloorApplied == false, "Explicit w:line list paragraphs should not report the default-auto floor.");
        TestAssert.Equal(1.2d, textLines[3].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(textLines[3].LineSpacingFactorFloorApplied == false, "Above-minimum default-auto list paragraphs should keep their factor without reporting the floor.");
        TestAssert.Equal(1.15d, textLines[4].EffectiveLineSpacingFactor ?? 0d);
        TestAssert.True(textLines[4].LineSpacingFactorFloorApplied == false, "Non-list paragraphs should not report the list floor.");
    }

    public static void DocxLayoutReservesListLabelFirstLineExtraLeading()
    {
        DocxListLabel tallLabel = new DocxListLabel(
            "*",
            "bullet",
            "*",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, "Label Metrics", new DocxRunFonts("Label Metrics", null, null, null, null, null, null, null)));
        DocxListLabel plainLabel = new DocxListLabel(
            "*",
            "bullet",
            "*",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, null, new DocxRunFonts(null, null, null, null, null, null, null, null)));
        DocxParagraph plainListHead = new DocxParagraph(
            [new DocxTextRun("Head", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.15d,
            null,
            new DocxParagraphSpacing(null, null, null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            plainLabel);
        DocxParagraph tallList = new DocxParagraph(
            [new DocxTextRun("Tall", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.15d,
            null,
            new DocxParagraphSpacing(null, null, null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            tallLabel);
        DocxParagraph plainList = new DocxParagraph(
            [new DocxTextRun("Flat", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            6d,
            0d,
            1.15d,
            null,
            new DocxParagraphSpacing(null, null, null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            plainLabel);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(plainListHead), new DocxParagraphElement(tallList), new DocxParagraphElement(plainList)],
            []);

        DocxLayoutItemSnapshot[] textLines = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None))
            .Pages[0]
            .Items
            .Where(item => item.Kind == "TextLine")
            .ToArray();

        // FamilyWidthTextMeasurer reports ascender 1.1x for the Label Metrics family and 1.0x otherwise,
        // so the tall-label item reserves 1pt of extra top leading (Office style-line probe 2026-09-06:
        // Symbol-bullet first-line gaps exceed body prediction while Calibri-bullet gaps match it).
        TestAssert.Equal(3, textLines.Length);
        double firstPitch = Math.Abs(textLines[1].Y - textLines[0].Y);
        double secondPitch = Math.Abs(textLines[2].Y - textLines[1].Y);
        TestAssert.True(Math.Abs(firstPitch - secondPitch - 1.0d) < 0.0001d, "The tall-label first line should reserve the label-ascender excess as extra top leading.");
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxLayoutItemSnapshot line = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None))
            .Pages[0]
            .Items
            .Single(item => item.Kind == "TextLine");

        TestAssert.Equal(10d, line.SingleLineHeightPoints ?? 0d);
        TestAssert.Equal(10d, line.ListLabelSingleLineHeightPoints ?? 0d);
        TestAssert.Equal(12d, line.BodyWindowsLineHeightPoints ?? 0d);
        TestAssert.Equal(14d, line.ListLabelWindowsLineHeightPoints ?? 0d);
        // Explicit w:line 276 (1.15) is honored as-authored for lists (Office explicit-115 probe 2026-09-06).
        TestAssert.Equal(11.5d, Math.Round(line.LineHeightPoints ?? 0d, 2));
        TestAssert.Equal("BodySingleLineAuto", line.LineHeightSource ?? string.Empty);
        TestAssert.True((line.ListLabelWindowsLineHeightPoints ?? 0d) > (line.BodyWindowsLineHeightPoints ?? 0d), "Snapshot should expose when list-label Windows extents exceed body extents.");
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
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(1, document.BodyElements.OfType<DocxManualBreakElement>().Count(element => element.Value == "column"));
    }

    public static void DocxMarkupFinalViewMergesDeletedParagraphMarksAndKeepsSpans()
    {
        string input = DocxTests.WriteDeletedParagraphMarkMarkupProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument simpleDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.SimpleMarkup);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("FirstSecond link field [c]|Visible list item", DocxTests.ParagraphTexts(finalDocument));
        TestAssert.Equal("FirstSecond link field [c]|Visible list item", DocxTests.ParagraphTexts(simpleDocument));
        TestAssert.Equal("First|Second link field [c]|Deleted list item|Visible list item", DocxTests.ParagraphTexts(originalDocument));
        TestAssert.Equal("First|Second link field [c]|Deleted list item|Visible list item", DocxTests.ParagraphTexts(allDocument));
        TestAssert.True(originalDocument.Paragraphs[0].HasDeletedParagraphMark && allDocument.Paragraphs[0].HasDeletedParagraphMark, "Original and all-markup views should preserve the deleted paragraph mark as a visible review artifact.");
        TestAssert.True(!finalDocument.Paragraphs[0].HasDeletedParagraphMark && !simpleDocument.Paragraphs[0].HasDeletedParagraphMark, "Final-style views should consume the deleted paragraph mark after merging the following paragraph.");
        TestAssert.Equal("1.", finalDocument.Paragraphs[1].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("1.", originalDocument.Paragraphs[2].ListLabel?.Text ?? string.Empty);
        TestAssert.Equal("2.", originalDocument.Paragraphs[3].ListLabel?.Text ?? string.Empty);

        DocxParagraph merged = finalDocument.Paragraphs[0];
        TestAssert.Equal(1, merged.Hyperlinks.Count);
        TestAssert.Equal(1, merged.FieldReferences.Count);
        TestAssert.Equal(1, merged.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment));
        TestAssert.True(merged.Hyperlinks.Single().TextRunStartIndex > 0 && merged.Hyperlinks.Single().SourceRunStartIndex > 0, "Merged hyperlinks should be shifted after the first paragraph's runs.");
        TestAssert.True(merged.FieldReferences.Single().TextRunIndex > 0 && merged.FieldReferences.Single().SourceRunIndex > 0, "Merged field references should be shifted after the first paragraph's runs.");
        TestAssert.True(merged.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Comment).SourceRunIndex > 0, "Merged comment anchors should keep a source run after index shifting.");
        TestAssert.True(new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(simpleDocument).RevisionItemCount >= 1, "Simple markup should still expose a change bar candidate for the consumed deleted paragraph mark.");
    }

    public static void DocxMarkupModesCollapseRevisedParagraphSpacingBeforeLayout()
    {
        string input = DocxTests.WriteRevisedParagraphSpacingCollapseProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("Before|Inserted spacing|After", DocxTests.ParagraphTexts(finalDocument));
        TestAssert.Equal("Before|Deleted spacing|After", DocxTests.ParagraphTexts(originalDocument));
        TestAssert.Equal("Before|Inserted spacing|Deleted spacing|After", DocxTests.ParagraphTexts(allDocument));

        AssertRevisedSpacing(finalDocument, OoxPdfDocxMarkupMode.Final, expectedLineCount: 3, revisedLineIndexes: [1], afterLineIndex: 2);
        AssertRevisedSpacing(originalDocument, OoxPdfDocxMarkupMode.Original, expectedLineCount: 3, revisedLineIndexes: [1], afterLineIndex: 2);
        AssertRevisedSpacing(allDocument, OoxPdfDocxMarkupMode.AllMarkup, expectedLineCount: 4, revisedLineIndexes: [1, 2], afterLineIndex: 3);

        DocxDocument deletedMarkFinal = DocxTests.ReadDocx(DocxTests.WriteDeletedParagraphMarkMarkupProbeDocx(), OoxPdfDocxMarkupMode.Final);
        DocxLayoutItemSnapshot[] deletedMarkLines = TextLineSnapshots(deletedMarkFinal, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(2, deletedMarkLines.Length);
        TestAssert.Equal(6d, deletedMarkLines[0].ParagraphAfterSpacingPoints ?? -1d);
        TestAssert.Equal(6d, deletedMarkLines[1].PendingAfterSpacingPoints ?? -1d);
        TestAssert.Equal(6d, deletedMarkLines[1].AppliedBeforeSpacingPoints ?? -1d);

        static void AssertRevisedSpacing(
            DocxDocument document,
            OoxPdfDocxMarkupMode mode,
            int expectedLineCount,
            IReadOnlyList<int> revisedLineIndexes,
            int afterLineIndex)
        {
            DocxLayoutItemSnapshot[] lines = TextLineSnapshots(document, mode);
            TestAssert.Equal(expectedLineCount, lines.Length);
            foreach (int lineIndex in revisedLineIndexes)
            {
                double expectedPending = lineIndex == revisedLineIndexes[0] ? 24d : 6d;
                TestAssert.Equal(expectedPending, lines[lineIndex].PendingAfterSpacingPoints ?? -1d);
                TestAssert.Equal(24d, lines[lineIndex].ParagraphBeforeSpacingPoints ?? -1d);
                TestAssert.Equal(0d, lines[lineIndex].AppliedBeforeSpacingPoints ?? -1d);
                TestAssert.Equal(6d, lines[lineIndex].ParagraphAfterSpacingPoints ?? -1d);
                TestAssert.True(lines[lineIndex].ContextualSpacingSuppressed == true, "Revised same-style paragraphs should suppress the before gap when contextual spacing is selected.");
            }

            TestAssert.Equal(6d, lines[afterLineIndex].PendingAfterSpacingPoints ?? -1d);
            TestAssert.Equal(6d, lines[afterLineIndex].AppliedBeforeSpacingPoints ?? -1d);
        }

        static DocxLayoutItemSnapshot[] TextLineSnapshots(DocxDocument document, OoxPdfDocxMarkupMode mode)
        {
            return new DocxRenderer(null, mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .InspectLayout(document)
                .Pages.SelectMany(page => page.Items)
                .Where(item => item.Kind == "TextLine")
                .ToArray();
        }
    }

    public static void DocxMarkupModesFilterBlockRevisionParagraphs()
    {
        string input = DocxTests.WriteBlockRevisionParagraphProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument simpleDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.SimpleMarkup);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("Before|Inserted paragraph|Moved to paragraph|After", DocxTests.ParagraphTexts(finalDocument));
        TestAssert.Equal("Before|Deleted paragraph|Moved from paragraph|After", DocxTests.ParagraphTexts(originalDocument));
        TestAssert.Equal("Before|Inserted paragraph|Moved to paragraph|After", DocxTests.ParagraphTexts(simpleDocument));
        TestAssert.Equal("Before|Inserted paragraph|Deleted paragraph|Moved from paragraph|Moved to paragraph|After", DocxTests.ParagraphTexts(allDocument));
        TestAssert.True(allDocument.Paragraphs.Any(paragraph => paragraph.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion)), "Block-level deletion revisions should be retained on paragraphs that survive the selected view.");
        TestAssert.True(allDocument.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Revision?.Kind == DocxRevisionKind.MoveFrom)), "Inherited block-level move revisions should reach paragraph runs.");
    }

    public static void DocxWordCompatibleAllMarkupScalesKeepNextParagraphEstimate()
    {
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 53d);
        DocxParagraph keep = new(
            [new DocxTextRun("Keep", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty with { KeepNext = true },
            null);
        DocxParagraph next = DocxTests.CreateDocxLayoutParagraph("Next", 10d, 10d);
        DocxDocument document = new(
            300d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(filler), new DocxParagraphElement(keep), new DocxParagraphElement(next)],
            [filler, keep, next],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxLayout reserve = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayout wordCompatible = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] wordLines = wordCompatible.Pages[0].Items.OfType<DocxTextLineLayout>().ToArray();

        TestAssert.Equal(2, reserve.Pages.Count);
        TestAssert.Equal(1, wordCompatible.Pages.Count);
        TestAssert.Equal(3, wordLines.Length);
        TestAssert.Equal("Filler", wordLines[0].Text);
        TestAssert.Equal("Keep", wordLines[1].Text);
        TestAssert.Equal("Next", wordLines[2].Text);
    }

    public static void DocxAllMarkupStylesTrackedChangeRuns()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();

        DocxParagraph paragraph = DocxTests.ReadTrackedChangeModeProbe(input, OoxPdfDocxMarkupMode.AllMarkup);

        DocxTextRun inserted = paragraph.Runs.Single(run => run.Text.Trim() == "Inserted");
        DocxTextRun deleted = paragraph.Runs.Single(run => run.Text.Trim() == "Deleted");
        DocxTextRun movedFrom = paragraph.Runs.Single(run => run.Text.Trim() == "MovedFrom");
        DocxTextRun movedTo = paragraph.Runs.Single(run => run.Text.Trim() == "MovedTo");
        TestAssert.True(inserted.Revision?.Kind == DocxRevisionKind.Insertion && inserted.Underline && inserted.ColorHex == "0000FF", "All-markup insertions should carry revision provenance and inline inserted styling.");
        TestAssert.True(deleted.Revision?.Kind == DocxRevisionKind.Deletion && deleted.Strike && deleted.ColorHex == "C00000", "All-markup deletions should carry revision provenance and inline deleted styling.");
        TestAssert.True(movedFrom.Revision?.Kind == DocxRevisionKind.MoveFrom && movedFrom.DoubleStrike && movedFrom.ColorHex == "C00000", "All-markup moved-from text should carry revision provenance and inline moved-from styling.");
        TestAssert.True(movedTo.Revision?.Kind == DocxRevisionKind.MoveTo && movedTo.Underline && movedTo.ColorHex == "008000", "All-markup moved-to text should carry revision provenance and inline moved-to styling.");
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
                      <w:spacing w:line="200" w:lineRule="exact"/>
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

        using FileStream stream = File.OpenRead(input);
        DocxDocument styledDocument = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
        ];
        DocxParagraph[] styledParagraphs = styledDocument.Paragraphs.ToArray();
        DocxBodyElement[] body = fillers.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>()
            .Concat(styledParagraphs.Select(paragraph => new DocxParagraphElement(paragraph)))
            .ToArray();
        var compactDocument = new DocxDocument(
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(compactDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] secondPageLines = layout.Pages[1].Items.OfType<DocxTextLineLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(4, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal("One", secondPageLines[0].Text);
        TestAssert.Equal("Two", secondPageLines[1].Text);
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
                      <w:spacing w:beforeAutospacing="1" w:afterAutospacing="1" w:beforeLines="120" w:afterLines="240" w:line="260" w:lineRule="atLeast"/>
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

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxStructureStyleUsageSnapshot styleUsage = DocxStructureSnapshot.FromDocument(document).StyleUsages.Single(usage => usage.Kind == "Paragraph");
        TestAssert.Equal(2, styleUsage.BeforeSpacingTokenParagraphCount);
        TestAssert.Equal(2, styleUsage.AfterSpacingTokenParagraphCount);
        TestAssert.Equal(2, styleUsage.BeforeAutoSpacingParagraphCount);
        TestAssert.Equal(2, styleUsage.AfterAutoSpacingParagraphCount);
        TestAssert.Equal(2, styleUsage.BeforeLinesSpacingParagraphCount);
        TestAssert.Equal(2, styleUsage.AfterLinesSpacingParagraphCount);
        TestAssert.Equal(2, styleUsage.ContextualSpacingParagraphCount);
        TestAssert.Equal(2, styleUsage.AtLeastLineSpacingParagraphCount);
    }
}
