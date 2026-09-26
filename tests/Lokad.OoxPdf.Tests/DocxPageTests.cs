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

internal static class DocxPageTests
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
                      <w:pgMar w:top="720" w:right="1440" w:bottom="1080" w:left="1800" w:header="1440" w:footer="1080" w:gutter="360"/>
                      <w:docGrid w:linePitch="326"/>
                      <w:footnotePr><w:pos w:val="beneathText"/><w:numStart w:val="3"/></w:footnotePr>
                      <w:endnotePr><w:pos w:val="sectEnd"/><w:numRestart w:val="eachSect"/></w:endnotePr>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
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
        TestAssert.Equal("360", settings.GutterDistanceValue ?? string.Empty);
        TestAssert.Equal(18d, settings.GutterDistancePoints ?? 0d);
        TestAssert.Equal("326", settings.DocGridLinePitchValue ?? string.Empty);
        TestAssert.Equal(16.3d, settings.DocGridLinePitchPoints ?? 0d);
        TestAssert.Equal("beneathText", settings.FootnoteReferenceSettings.PositionValue ?? string.Empty);
        TestAssert.Equal(3, settings.FootnoteReferenceSettings.NumberStart ?? 0);
        TestAssert.Equal("sectEnd", settings.EndnoteReferenceSettings.PositionValue ?? string.Empty);
        TestAssert.Equal("eachSect", settings.EndnoteReferenceSettings.NumberRestartValue ?? string.Empty);
        DocxLayoutPageSnapshot pageSnapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document).Pages.Single();
        TestAssert.Equal(108d, pageSnapshot.MarginLeft);
        TestAssert.Equal(72d, pageSnapshot.MarginRight);
        TestAssert.Equal("beneathText", pageSnapshot.SectionFootnotePositionValue ?? string.Empty);
        TestAssert.Equal("sectEnd", pageSnapshot.SectionEndnotePositionValue ?? string.Empty);
        TestAssert.Equal("eachSect", pageSnapshot.SectionEndnoteNumberRestartValue ?? string.Empty);
        TestAssert.Equal(842d, document.PageWidthPoints);
        TestAssert.Equal(595d, document.PageHeightPoints);
        TestAssert.Equal(90d, document.MarginLeftPoints);
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
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d);
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

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

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
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d);
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

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/footer", annotation.Uri);
        TestAssert.True(annotation.Y < 60d, "Footer annotation should be anchored near the page bottom static story.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover footer hyperlink text.");
    }

    public static void DocxFontPlanIncludesAllHeaderFooterVariants()
    {
        var defaultHeader = DocxTests.CreateFontPlanParagraph(new DocxTextRun("DefaultHeader", 11d, null, false, false, false, null, "Default Sans")
        {
            Fonts = new DocxRunFonts("Default Sans", null, null, null, null, null, null, null)
        });
        var firstHeader = DocxTests.CreateFontPlanParagraph(new DocxTextRun("FirstHeader", 11d, null, false, false, false, null, "First Sans")
        {
            Fonts = new DocxRunFonts("First Sans", null, null, null, null, null, null, null)
        });
        var evenFooter = DocxTests.CreateFontPlanParagraph(new DocxTextRun("EvenFooter", 11d, null, false, false, false, null, "Even Sans")
        {
            Fonts = new DocxRunFonts("Even Sans", null, null, null, null, null, null, null)
        });
        var sectionHeader = DocxTests.CreateFontPlanParagraph(new DocxTextRun("SectionHeader", 11d, null, false, false, false, null, "Section Sans")
        {
            Fonts = new DocxRunFonts("Section Sans", null, null, null, null, null, null, null)
        });
        var sectionFooter = DocxTests.CreateFontPlanParagraph(new DocxTextRun("SectionFooter", 11d, null, false, false, false, null, "Section Footer Sans")
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
            [new DocxSectionBreakElement(sectionSettings, DocxSectionBreakType.NextPage, null, null, null, [])],
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

        string plannedTexts = string.Join("|", DocxFontPlan.Create(document, resolver, CancellationToken.None).Runs.Select(run => run.Run.Text).Order(StringComparer.Ordinal));

        TestAssert.Equal("DefaultHeader|EvenFooter|FirstHeader|SectionFooter|SectionHeader", plannedTexts);
    }

    public static void DocxLayoutUsesResolvedBodyPageFieldForLineBreaking()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("A {PAGE} B", 10d, 10d);
        var document = new DocxDocument(
            45d,
            100d,
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
        TestAssert.Equal("A 1 B", lines[0].Text);
    }

    public static void DocxLayoutUsesCompactBodyNumPagesFieldForLineBreaking()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("A {NUMPAGES} B", 10d, 10d);
        var document = new DocxDocument(
            45d,
            100d,
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
        TestAssert.Equal("A {NUMPAGES} B", lines[0].Text);
    }

    public static void DocxLayoutRightAlignsBodyNumPagesUsingCompactFieldWidth()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A {NUMPAGES} B", 10d, null, false, false, false, null, null) { FieldKind = DocxFieldKind.NumPages }],
            [],
            null,
            DocxTextAlignment.Right,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
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
            [new DocxParagraphElement(paragraph)],
            [],
            []);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.True(line.X > 60d, "Right alignment should use the compact dynamic-field width, not the literal NUMPAGES placeholder width.");
        TestAssert.True(line.Width < 30d, "Line width should reflect the compact layout-time NUMPAGES proxy.");
        TestAssert.Equal("A {NUMPAGES} B", line.Text);
    }

    public static void DocxLayoutRightAlignsFloatingTextBoxNumPagesUsingCompactFieldWidth()
    {
        var textBoxParagraph = new DocxParagraph(
            [new DocxTextRun("A {NUMPAGES} B", 10d, null, false, false, false, null, null) { FieldKind = DocxFieldKind.NumPages }],
            [],
            null,
            DocxTextAlignment.Right,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxDocument document = new(
            300d,
            300d,
            30d,
            90d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d))],
            [],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout line = layout.FloatingDrawings.Single().TextBoxLayout!.TextLines.Single();

        // Inset-narrowed content shifts the right-aligned line left by the left inset.
        TestAssert.True(line.X > 70d - 7.2d, "Floating text-box right alignment should use the compact dynamic-field width, not the literal NUMPAGES placeholder width.");
        TestAssert.True(line.Width < 30d, "Floating text-box line width should reflect the compact layout-time NUMPAGES proxy.");
        TestAssert.Equal("A {NUMPAGES} B", line.Text);
    }

    public static void DocxTextEmissionResolvesFloatingTextBoxPageFieldsAtEmission()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        var textBoxParagraph = new DocxParagraph(
            [
                new DocxTextRun("Box ", 10d, null, false, false, false, null, familyName),
                new DocxTextRun("{PAGE}", 10d, null, false, false, false, null, familyName) { FieldKind = DocxFieldKind.Page },
                new DocxTextRun(" of ", 10d, null, false, false, false, null, familyName),
                new DocxTextRun("{NUMPAGES}", 10d, null, false, false, false, null, familyName) { FieldKind = DocxFieldKind.NumPages }
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
            null);
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxParagraph firstPage = DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d);
        DocxParagraph secondPage = DocxTests.CreateDocxLayoutParagraph("Second page", 10d, 12d);
        var document = new DocxDocument(
            300d,
            300d,
            30d,
            90d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [new DocxParagraphElement(firstPage), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(secondPage)],
            [firstPage, secondPage],
            []);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionLineSnapshot textBoxLine = renderer.InspectTextEmission(document).Lines
            .Single(line => line.StoryKind == "TextBox");
        DocxTextEmissionSegmentSnapshot[] visibleSegments = textBoxLine.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();
        DocxTextEmissionSegmentSnapshot[] digitSegments = visibleSegments
            .Where(segment => segment.TextLength == 1 && segment.CharacterProfile.DigitCount == 1)
            .ToArray();

        TestAssert.Equal(0, textBoxLine.PageIndex);
        TestAssert.Equal("Body", textBoxLine.ContainerStoryKind ?? string.Empty);
        TestAssert.Equal(10, visibleSegments.Sum(segment => segment.TextLength));
        TestAssert.Equal(2, digitSegments.Length);
        TestAssert.True(Math.Abs(digitSegments[1].Width - digitSegments[1].AdvanceProfile.NaturalPdfWidth) < 0.5d, "Resolved floating text-box NUMPAGES should use the emitted digit advance instead of the literal placeholder advance.");
    }

    public static void DocxLayoutUsesResolvedPageFieldForKeepEstimate()
    {
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Fill", 20d, 20d);
        DocxParagraph kept = DocxTests.CreateDocxLayoutParagraph(
            "A {PAGE} B",
            10d,
            10d,
            keepRules: new DocxParagraphKeepRules(null, null, true, null, null, null));
        var document = new DocxDocument(
            45d,
            60d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(filler), new DocxParagraphElement(kept)],
            [],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(1, layout.Pages.Count);
        DocxTextLineLayout keptLine = layout.Pages[0].Items
            .OfType<DocxTextLineLayout>()
            .Single(line => line.SourceParagraph == kept);
        TestAssert.Equal("A 1 B", keptLine.Text);
    }

    public static void DocxTextEmissionCompactsResolvedNumPagesFieldAdvance()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        var paragraph = new DocxParagraph(
            [
                new DocxTextRun("A", 10d, null, false, false, false, null, familyName),
                new DocxTextRun("{NUMPAGES}", 10d, null, false, false, false, null, familyName) { FieldKind = DocxFieldKind.NumPages },
                new DocxTextRun("B", 10d, null, false, false, false, null, familyName)
            ],
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
        DocxParagraph secondPage = DocxTests.CreateDocxLayoutParagraph("Second", 10d, 10d);
        var document = new DocxDocument(
            200d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(paragraph),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(secondPage)
            ],
            [paragraph, secondPage],
            []);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionSegmentSnapshot[] visibleSegments = renderer.InspectTextEmission(document)
            .Lines
            .First(line => line.SourceBlockIndex == 0)
            .Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();

        TestAssert.Equal(3, visibleSegments.Length);
        TestAssert.Equal(1, visibleSegments[1].TextLength);
        TestAssert.True(Math.Abs(visibleSegments[1].Width - visibleSegments[1].AdvanceProfile.NaturalPdfWidth) < 0.5d, "Resolved NUMPAGES should use the emitted digit advance instead of the literal placeholder advance.");
        TestAssert.True(Math.Abs(visibleSegments[2].X - (visibleSegments[1].X + visibleSegments[1].Width)) < 0.5d, "Runs after NUMPAGES should start after the compact emitted field advance.");
    }

    public static void DocxTextEmissionResolvesPlacedFootnotePageFieldsAtEmission()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "9",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        var footnoteParagraph = new DocxParagraph(
            [
                new DocxTextRun("Note ", 10d, null, false, false, false, null, familyName),
                new DocxTextRun("{PAGE}", 10d, null, false, false, false, null, familyName) { FieldKind = DocxFieldKind.Page },
                new DocxTextRun(" of ", 10d, null, false, false, false, null, familyName),
                new DocxTextRun("{NUMPAGES}", 10d, null, false, false, false, null, familyName) { FieldKind = DocxFieldKind.NumPages }
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
            null);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "9",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        DocxParagraph secondPage = DocxTests.CreateDocxLayoutParagraph("Second page", 10d, 12d);
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
            [
                new DocxParagraphElement(bodyParagraph),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(secondPage)
            ],
            [bodyParagraph, secondPage],
            [])
        {
            RelatedStories = [footnoteStory]
        };
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionLineSnapshot footnoteLine = renderer.InspectTextEmission(document).Lines
            .Single(line => line.StoryKind == "Footnote");
        DocxTextEmissionSegmentSnapshot[] visibleSegments = footnoteLine.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();
        DocxTextEmissionSegmentSnapshot[] digitSegments = visibleSegments
            .Where(segment => segment.TextLength == 1 && segment.CharacterProfile.DigitCount == 1)
            .ToArray();

        TestAssert.Equal(0, footnoteLine.PageIndex);
        TestAssert.True(!footnoteLine.IsStaticStory, "Placed footnote lines should be emitted as page-owned related story content.");
        TestAssert.Equal(11, visibleSegments.Sum(segment => segment.TextLength));
        TestAssert.Equal(2, digitSegments.Length);
        TestAssert.True(digitSegments[0].CharacterProfile.OtherCount == 0, "PAGE should emit the owning page digit.");
        TestAssert.True(digitSegments[1].CharacterProfile.OtherCount == 0, "NUMPAGES should emit the final page-count digit for placed footnotes.");
        TestAssert.True(Math.Abs(digitSegments[1].Width - digitSegments[1].AdvanceProfile.NaturalPdfWidth) < 0.5d, "Resolved footnote NUMPAGES should use the emitted digit advance instead of the literal placeholder advance.");
    }

    public static void DocxTextEmissionResolvesPlacedEndnotePageFieldsAtEmission()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "11",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        var endnoteParagraph = new DocxParagraph(
            [
                new DocxTextRun("Note ", 10d, null, false, false, false, null, familyName),
                new DocxTextRun("{PAGE}", 10d, null, false, false, false, null, familyName) { FieldKind = DocxFieldKind.Page },
                new DocxTextRun(" of ", 10d, null, false, false, false, null, familyName),
                new DocxTextRun("{NUMPAGES}", 10d, null, false, false, false, null, familyName) { FieldKind = DocxFieldKind.NumPages }
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
            null);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "11",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxParagraph secondPage = DocxTests.CreateDocxLayoutParagraph("Second page", 10d, 12d);
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
            [
                new DocxParagraphElement(bodyParagraph),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(secondPage)
            ],
            [bodyParagraph, secondPage],
            [])
        {
            RelatedStories = [endnoteStory]
        };
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionLineSnapshot endnoteLine = renderer.InspectTextEmission(document).Lines
            .Single(line => line.StoryKind == "Endnote");
        DocxTextEmissionSegmentSnapshot[] visibleSegments = endnoteLine.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();
        DocxTextEmissionSegmentSnapshot[] digitSegments = visibleSegments
            .Where(segment => segment.TextLength == 1 && segment.CharacterProfile.DigitCount == 1)
            .ToArray();

        TestAssert.Equal(1, endnoteLine.PageIndex);
        TestAssert.True(!endnoteLine.IsStaticStory, "Placed endnote lines should be emitted as page-owned related story content.");
        TestAssert.Equal(11, visibleSegments.Sum(segment => segment.TextLength));
        TestAssert.Equal(2, digitSegments.Length);
        TestAssert.True(digitSegments[0].CharacterProfile.OtherCount == 0, "PAGE should emit the owning page digit.");
        TestAssert.True(digitSegments[1].CharacterProfile.OtherCount == 0, "NUMPAGES should emit the final page-count digit for placed endnotes.");
        TestAssert.True(Math.Abs(digitSegments[1].Width - digitSegments[1].AdvanceProfile.NaturalPdfWidth) < 0.5d, "Resolved endnote NUMPAGES should use the emitted digit advance instead of the literal placeholder advance.");
    }

    public static void DocxSyntheticParagraphsBreakAcrossPages()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxPageBreakElement[] breaks = document.BodyElements.OfType<DocxPageBreakElement>().ToArray();

        TestAssert.Equal(2, breaks.Length);
        TestAssert.Equal(DocxBreakSourceKind.PageBreakBefore, breaks[0].SourceKind);
        TestAssert.Equal("on", breaks[0].Value ?? string.Empty);
        TestAssert.Equal(DocxBreakSourceKind.PageBreakBefore, breaks[1].SourceKind);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxPageBreakElement[] breaks = document.BodyElements.OfType<DocxPageBreakElement>().ToArray();

        TestAssert.Equal(1, breaks.Length);
        TestAssert.Equal(DocxBreakSourceKind.PageBreakBefore, breaks[0].SourceKind);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxBodyElement[] elements = document.BodyElements.ToArray();

        TestAssert.Equal(3, elements.Length);
        TestAssert.True(elements[1] is DocxPageBreakElement, "A run-level page-break-only paragraph should become a body page break.");
        var pageBreak = (DocxPageBreakElement)elements[1];
        TestAssert.Equal(DocxBreakSourceKind.RunBreak, pageBreak.SourceKind);
        TestAssert.Equal("page", pageBreak.Value ?? string.Empty);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxBodyElement[] elements = document.BodyElements.ToArray();
        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("Alpha", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxPageBreakElement, "The inline page break should become a body page break.");
        TestAssert.Equal("Beta", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("Alpha", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal("Beta", layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().Text);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxBodyElement[] elements = document.BodyElements.ToArray();
        TestAssert.Equal(3, elements.Length);
        TestAssert.Equal("Alpha", ((DocxParagraphElement)elements[0]).Paragraph.Runs.Single().Text);
        TestAssert.True(elements[1] is DocxPageBreakElement, "The simple-field-contained page break should become a body page break.");
        TestAssert.Equal("Beta", ((DocxParagraphElement)elements[2]).Paragraph.Runs.Single().Text);
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
                new DocxSectionBreakElement(DocxPageSettings.Empty, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("First", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal("Second", layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().Text);
    }

    public static void DocxSectionBreakPageSettingsOwnPrecedingLayoutSection()
    {
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Second", fontSize: 10d, lineSpacingPoints: 10d);
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
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, "2", "1", "720", []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

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
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Second", fontSize: 10d, lineSpacingPoints: 10d);
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
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.Continuous, null, null, null, []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

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
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Second", fontSize: 10d, lineSpacingPoints: 10d);
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
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.OddPage, null, null, null, []),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

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
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 12d);
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
                    DocxSectionBreakType.NextPage,
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

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

    // RV15: repeated anchor lookups must not rescan all layout items per
    // drawing: one index build visits each item once, lookups visit nothing.
    public static void FloatingSourceBlockLookupsStayBounded()
    {
        DocxLayout layout = CreateFloatingSourceBlockDocumentLayout();
        TestAssert.True(layout.Pages.Count >= 2, "Fixture must span pages.");
        var counter = new ItemVisitCounter();
        DocxLayoutPage[] countingPages = layout.Pages.Select(page => page with { Items = new CountingItems(page.Items, counter) }).ToArray();
        int totalItems = countingPages.Sum(page => page.Items.Count);
        TestAssert.True(totalItems > 0, "Fixture must lay out items.");
        object index = InvokeBuildSourceBlockIndex(countingPages);
        TestAssert.True(counter.Visits <= 2L * totalItems, string.Format("Index build must visit each item once, visited {0} for {1} items.", counter.Visits, totalItems));
        long afterBuild = counter.Visits;
        for (int k = 0; k < 25; k++)
        {
            TryLookupBlock(index, k % 4, out _);
            TryLookupBlock(index, 999, out _);
        }
        TestAssert.Equal(afterBuild, counter.Visits);
    }

    // RV15: the per-layout source-block index must agree with the linear scan
    // slot by slot, including missing blocks.
    public static void FloatingSourceBlockIndexAgreesWithLinearScan()
    {
        DocxLayout layout = CreateFloatingSourceBlockDocumentLayout();
        IReadOnlyList<DocxLayoutPage> pages = layout.Pages;
        object index = InvokeBuildSourceBlockIndex(pages);
        int[] blocks = DocxLayoutSnapshot.FromLayout(layout).SourceBlocks.Select(block => block.SourceBlockIndex).ToArray();
        TestAssert.True(blocks.Length >= 5, "Fixture must lay out several blocks.");
        int low = blocks.Min() - 1;
        int high = blocks.Max() + 1;
        int found = 0;
        for (int block = low; block <= high; block++)
        {
            object? expected = InvokeFindSourceBlockBounds(pages, block);
            bool hit = TryLookupBlock(index, block, out object? actual);
            TestAssert.True((expected is null) == !hit, string.Format("Block {0} presence must match.", block));
            if (expected is not null)
            {
                TestAssert.True(expected.Equals(actual), string.Format("Block {0} bounds must match.", block));
                found++;
            }
        }
        TestAssert.True(found >= 5, "Agreement must cover several present blocks.");
    }

    private static DocxLayout CreateFloatingSourceBlockDocumentLayout()
    {
        var paragraphs = new List<DocxParagraph>();
        var elements = new List<DocxBodyElement>();
        DocxTable? table = null;
        for (int i = 0; i < 12; i++)
        {
            DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Block paragraph " + i, 10d, 12d);
            paragraphs.Add(paragraph);
            elements.Add(new DocxParagraphElement(paragraph));
            if (i == 5)
            {
                table = DocxTests.CreateSingleCellTable("Cell text", 12d);
                elements.Add(new DocxTableElement(table));
            }
        }

        int[] anchoredBlocks = [0, 1, 2, 0, 1, 2, 0, 1, 2, 999];
        var drawings = new List<DocxFloatingDrawing>();
        foreach (int block in anchoredBlocks)
        {
            drawings.Add(new DocxFloatingDrawing("0", "0", "0", "0", "0", "0", "0", "0", "1", "1", "914400", "457200", "column", "left", null, "paragraph", null, "0", DocxFloatingWrapKind.Square, "bothSides", SourceParagraphIndex: 0, SourceBlockIndex: block, ImageRelationshipId: null, Image: null));
        }
        var document = new DocxDocument(200d, 120d, 10d, 10d, 10d, 10d, DocxPageSettings.Empty, drawings, [], [], elements, paragraphs, table is null ? [] : [table]);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
    }

    private static object InvokeBuildSourceBlockIndex(object pages)
    {
        System.Reflection.MethodInfo build = typeof(DocxLayoutEngine).GetMethod("BuildSourceBlockIndex", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected source-block index.");
        try
        {
            return build.Invoke(null, [pages, CancellationToken.None])!;
        }
        catch (System.Reflection.TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object? InvokeFindSourceBlockBounds(object pages, int block)
    {
        System.Reflection.MethodInfo find = typeof(DocxLayoutEngine).GetMethod("FindSourceBlockBounds", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected source-block scan.");
        try
        {
            return find.Invoke(null, [pages, block]);
        }
        catch (System.Reflection.TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static bool TryLookupBlock(object index, int block, out object? value)
    {
        System.Reflection.MethodInfo lookup = index.GetType().GetMethod("TryGetValue")
            ?? throw new InvalidOperationException("Expected index lookup.");
        object?[] args = [block, null];
        bool found = (bool)lookup.Invoke(index, args)!;
        value = args[1];
        return found;
    }

    private sealed class ItemVisitCounter
    {
        public long Visits;
    }
    private sealed class CountingItems(IReadOnlyList<DocxLayoutItem> inner, ItemVisitCounter counter) : IReadOnlyList<DocxLayoutItem>
    {
        public DocxLayoutItem this[int index] => inner[index];
        public int Count => inner.Count;
        public IEnumerator<DocxLayoutItem> GetEnumerator()
        {
            foreach (DocxLayoutItem item in inner)
            {
                counter.Visits++;
                yield return item;
            }
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public static void DocxPageLayoutStageManualColumnBreakAdvancesActiveColumn()
    {
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 12d);
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
            DocxFloatingWrapKind.Square,
            "bothSides",
            SourceParagraphIndex: 0,
            SourceBlockIndex: 2, ImageRelationshipId: null, Image: null);
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
                new DocxManualBreakElement(DocxBreakSourceKind.RunBreak, "column", null),
                new DocxParagraphElement(second),
                new DocxSectionBreakElement(sectionSettings, DocxSectionBreakType.NextPage, "2", "1", "360", [])
            ],
            [first, second],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
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

    public static void DocxSyntheticParagraphKeepLinesStartsBlockOnNextPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        DocxParagraph kept = DocxTests.CreateDocxLayoutParagraph(
            "First Second",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(null, null, true, null, null, null));
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 15d, lineSpacingPoints: 15d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 15d, lineSpacingPoints: 15d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 15d, lineSpacingPoints: 15d)
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FillFirstSecond".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded, CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(3, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(2, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Count());
    }

    public static void DocxSyntheticParagraphKeepNextMovesPairToNextPage()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        DocxParagraph keepNext = DocxTests.CreateDocxLayoutParagraph(
            "Keep",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph next = DocxTests.CreateDocxLayoutParagraph("Next", fontSize: 11d, lineSpacingPoints: 11d);
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FillKeepNext".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded, CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(4, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(2, layout.Pages[1].Items.OfType<DocxTextLineLayout>().Count());
    }

    public static void DocxSyntheticPageKeepNextEstimatesIndentedBlockTarget()
    {
        DocxParagraph keepNext = DocxTests.CreateDocxLayoutParagraph(
            "Keep",
            fontSize: 10d,
            lineSpacingPoints: 10d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph tableParagraph = DocxTests.CreateDocxLayoutParagraph("One Two Three Four Five", fontSize: 10d, lineSpacingPoints: 10d);
        var table = new DocxTable(
            null,
            [100d],
            [new DocxTableRow([new DocxTableCell("One Two Three Four Five", [tableParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 10d)]) with {PreferredWidthPoints = 100d,IndentPoints = 70d };
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        DocxParagraph keepFirst = DocxTests.CreateDocxLayoutParagraph(
            "Keep",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph keepSecond = DocxTests.CreateDocxLayoutParagraph(
            "Chain",
            fontSize: 11d,
            lineSpacingPoints: 11d,
            keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        DocxParagraph end = DocxTests.CreateDocxLayoutParagraph("End", fontSize: 11d, lineSpacingPoints: 11d);
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 9d, lineSpacingPoints: 9d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 9d, lineSpacingPoints: 9d)
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FillKeepChainEnd".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded, CancellationToken.None);
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        DocxParagraph widowControlled = DocxTests.CreateDocxLayoutParagraph(
            "One\nTwo\nThree",
            fontSize: 10d,
            lineSpacingPoints: 10d,
            keepRules: new DocxParagraphKeepRules(null, null, null, null, true, null));
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FillOneTwoThree".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded, CancellationToken.None);
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        DocxParagraph widowControlled = DocxTests.CreateDocxLayoutParagraph("One\nTwo\nThree", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph[] fillers =
        [
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d),
            DocxTests.CreateDocxLayoutParagraph("Fill", fontSize: 10d, lineSpacingPoints: 10d)
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "FillOneTwoThree".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTextLineLayout[] secondPageLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[1]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(3, secondPageLines.Length);
        TestAssert.Equal("One", secondPageLines[0].Text);
    }

    public static void DocxBreakOnlyParagraphEmitsOfficeSpacingRows()
    {
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("Break spill   ", 10d, 12d);
        DocxParagraph breakParagraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 12d);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Next", 10d, 12d);
        var document = new DocxDocument(
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
            [
                new DocxParagraphElement(first),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", breakParagraph),
                new DocxParagraphElement(second)
            ],
            [first, second],
            []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(2, layout.Pages.Count);
        DocxTextLineLayout[] firstPageLines = layout.Pages[0].Items.OfType<DocxTextLineLayout>().ToArray();
        TestAssert.Equal(2, firstPageLines.Length);
        TestAssert.Equal("Break spill    ", firstPageLines[0].Text);
        TestAssert.Equal("  ", firstPageLines[1].Text);
        TestAssert.Equal(2, firstPageLines[1].Segments.Count);
        TestAssert.True(Math.Abs((firstPageLines[1].Segments[1].X - firstPageLines[1].Segments[0].X) - 144d) < 1d, "spill tab stop");
        DocxParagraph control = DocxTests.CreateDocxLayoutParagraph("Control", 10d, 12d);
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 12d);
        var controlDocument = new DocxDocument(
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
            [new DocxParagraphElement(control), new DocxParagraphElement(after)],
            [control, after],
            []);
        DocxTextLineLayout[] controlLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(controlDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();
        TestAssert.Equal(2, controlLines.Length);
        TestAssert.True(controlLines[0].Text.EndsWith(" ", StringComparison.Ordinal) == false, "no spill without break");
        DocxParagraph tail = DocxTests.CreateDocxLayoutParagraph("Tail   ", 10d, 12d);
        DocxParagraph tailAfter = DocxTests.CreateDocxLayoutParagraph("After", 10d, 12d);
        var tailDocument = new DocxDocument(
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
            [new DocxParagraphElement(tail), new DocxParagraphElement(tailAfter)],
            [tail, tailAfter],
            []);
        DocxTextLineLayout[] tailLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(tailDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();
        TestAssert.Equal(2, tailLines.Length);
        TestAssert.Equal("Tail    ", tailLines[0].Text);
    }

    public static void DocxAlignedTrailingLinesShareDrawableLinePositions()
    {
        // RV06 align matrix: Office centers/rights drawable text, letting authored
        // and added trailing spaces overflow past the edge.
        DocxParagraph centerTrail = DocxTests.CreateDocxLayoutParagraph("C trail   ", 10d, 12d) with { Alignment = DocxTextAlignment.Center };
        DocxParagraph centerClean = DocxTests.CreateDocxLayoutParagraph("C trail", 10d, 12d) with { Alignment = DocxTextAlignment.Center };
        DocxParagraph rightTrail = DocxTests.CreateDocxLayoutParagraph("R trail   ", 10d, 12d) with { Alignment = DocxTextAlignment.Right };
        DocxParagraph rightClean = DocxTests.CreateDocxLayoutParagraph("R trail", 10d, 12d) with { Alignment = DocxTextAlignment.Right };
        var document = new DocxDocument(
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
            [
                new DocxParagraphElement(centerTrail),
                new DocxParagraphElement(centerClean),
                new DocxParagraphElement(rightTrail),
                new DocxParagraphElement(rightClean)
            ],
            [centerTrail, centerClean, rightTrail, rightClean],
            []);
        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();
        TestAssert.Equal(4, lines.Length);
        TestAssert.Equal(lines[1].X, lines[0].X);
        TestAssert.Equal(lines[3].X, lines[2].X);
    }

    public static void DocxBodyExactFirstBaselineFollowsOfficeRatio()
    {
        // RV06 pagination probe (edge-page-ex48-body, Word 16.0): Office drops the
        // first baseline of exact-spaced body text to 0.8 x the exact line height
        // below the content top (38.4 at exact-48), while the renderer applies the
        // 0.299em bottom inset (44.41 at 12pt). Pre-fix the baseline sits 6pt deep.
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Probe", 12d, 48d) with
        {
            Spacing = new DocxParagraphSpacing(null, null, null, null, null, null, null, "exact", null)
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        double baselineY = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single().Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        TestAssert.True(
            Math.Abs(baselineY - (190d - 38.4d)) < 0.001d,
            $"Exact-48pt body first baseline should sit 38.4pt below the content top. baselineY={baselineY}.");
    }
}
