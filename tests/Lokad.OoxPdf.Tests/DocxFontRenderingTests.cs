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

internal static class DocxFontRenderingTests
{
    public static void DocxStructureSnapshotUsesBodyElementInventoryAsCanonicalSource()
    {
        DocxParagraph cellParagraph = DocxTests.CreateDocxLayoutParagraph("Cell", fontSize: 9d, lineSpacingPoints: 10d) with
        {
            StyleId = "CellStyle"
        };
        var table = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([new DocxTableCell(string.Empty, [cellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], null)]) with {StyleId = "TableStyle" };
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

    public static void DocxFontPlanTextMeasurerUsesResolvedFontFace()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string text = "Office metrics";
        var run = new DocxTextRun(text, 12d, null, font.Value.Resolution.Bold, font.Value.Resolution.Italic, false, null, font.Value.Resolution.FamilyName)
        {
            Fonts = new DocxRunFonts(font.Value.Resolution.FamilyName, null, null, null, null, null, null, null)
        };
        DocxDocument document = DocxTests.CreateFontPlanDocument(run, new DocxFontCatalog([], DocxThemeFonts.Empty));
        var resolver = new DocxTests.SingleResolutionFontResolver(font.Value.Resolution);
        DocxFontPlan plan = DocxFontPlan.Create(document, resolver, CancellationToken.None);

        double measured = new DocxFontPlanTextMeasurer(plan, null, CancellationToken.None).MeasureText(run, text, run.FontSize);
        double expected = DocxTests.MeasureOpenTypeText(font.Value.Font, text, run.FontSize);

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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .Single();

        TestAssert.Equal(2, line.Segments.Count);
        TestAssert.Equal(5d, line.Segments[0].Width);
        TestAssert.Equal(40d, line.Segments[1].Width);
        TestAssert.Equal(line.Segments[0].X + line.Segments[0].Width, line.Segments[1].X);
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

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
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

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal("A ", lines[0].Text);
        TestAssert.Equal("B", lines[1].Text);
        TestAssert.Equal("Wide", lines[1].Segments.Single().StyleRun.FontFamily ?? string.Empty);
    }

    public static void DocxRendererUsesThemeTypefaceForThemeOnlyDefaultRun()
    {
        var resolver = new WindowsFontResolver(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"));
        string defaultFamily = resolver.Resolve(new FontRequest(DocxRenderer.DefaultDocumentTypefaceRequest)).FamilyName;
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFontExcept(defaultFamily);
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
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
        TestAssert.Contains("/BaseFont /", pdf); TestAssert.Contains("+" + PdfEmbeddedFont.SanitizeName(font.Value.Resolution.FamilyName), pdf);
    }

    public static void DocxRendererUsesFontCatalogAlternateBeforeResolverFallback()
    {
        var resolver = new WindowsFontResolver(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"));
        string defaultFamily = resolver.Resolve(new FontRequest(DocxRenderer.DefaultDocumentTypefaceRequest)).FamilyName;
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFontExcept(defaultFamily);
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
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
        TestAssert.Contains("/BaseFont /", pdf); TestAssert.Contains("+" + PdfEmbeddedFont.SanitizeName(font.Value.Resolution.FamilyName), pdf);
    }

    public static void DocxRendererEmbedsResolvedTrueTypeCollectionFace()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(fontsDirectory))");
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        (FontFaceResolution Resolution, string FirstFamily)? collectionFace = resolver.GetDiscoveredFonts()
            .Where(f => f.FontFaceIndex > 0 && f.Source is FileFontProgramSource && Regex.IsMatch(f.FamilyName, @"^[A-Za-z0-9 ._-]+$"))
            .Select(f => DocxTests.TryLoadCollectionFace(f))
            .Where(item => item is not null)
            .Select(item => item!.Value)
            .FirstOrDefault(item => !item.Resolution.FamilyName.Equals(item.FirstFamily, StringComparison.OrdinalIgnoreCase) &&
                !PdfEmbeddedFont.SanitizeName(item.Resolution.FamilyName).StartsWith(PdfEmbeddedFont.SanitizeName(item.FirstFamily), StringComparison.Ordinal) &&
                !PdfEmbeddedFont.SanitizeName(item.FirstFamily).StartsWith(PdfEmbeddedFont.SanitizeName(item.Resolution.FamilyName), StringComparison.Ordinal));
        if (collectionFace is null)
        {
            TestAssert.Skip("Environmental precondition not met: (collectionFace is null)");
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
        string expected = "+" + PdfEmbeddedFont.SanitizeName(collectionFace.Value.Resolution.FamilyName);
        string firstFace = "+" + PdfEmbeddedFont.SanitizeName(collectionFace.Value.FirstFamily);
        TestAssert.Contains("/BaseFont /", pdf); TestAssert.Contains(expected, pdf);
        TestAssert.True(!pdf.Contains(firstFace, StringComparison.Ordinal), "Expected DOCX embedding to honor the resolved TrueType collection face index.");
    }

    public static void DocxRendererEmitsDistinctResourcesForResolvedRunTypefaces()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!Directory.Exists(fontsDirectory))
        {
            TestAssert.Skip("Environmental precondition not met: (!Directory.Exists(fontsDirectory))");
        }

        (FontFaceResolution Resolution, OpenTypeFont Font)? first = DocxTests.FindUsableInstalledFont();
        if (first is null)
        {
            TestAssert.Skip("Environmental precondition not met: (first is null)");
        }

        (FontFaceResolution Resolution, OpenTypeFont Font)? second = DocxTests.FindUsableInstalledFontExcept(first.Value.Resolution.FamilyName);
        if (second is null ||
            string.Equals(first.Value.Resolution.Source.StableId, second.Value.Resolution.Source.StableId, StringComparison.OrdinalIgnoreCase) &&
            first.Value.Resolution.FontFaceIndex == second.Value.Resolution.FontFaceIndex)
        {
            TestAssert.Skip("Environmental precondition not met: (second is null || string.Equals(first.Value.Resolution.Source.StableId, second.Value.Resolution.Source.StableId, StringComparison.OrdinalIgnoreCase) && first.Value.Resolution.FontFaceIndex == second.Value.Resolution.FontFaceIndex)");
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
        TestAssert.Contains("/BaseFont /", pdf); TestAssert.Contains("+" + PdfEmbeddedFont.SanitizeName(first.Value.Resolution.FamilyName), pdf);
        TestAssert.Contains("/BaseFont /", pdf); TestAssert.Contains("+" + PdfEmbeddedFont.SanitizeName(second.Value.Resolution.FamilyName), pdf);
    }

    public static void DocxRendererDoesNotSynthesizeBoldWhenResolvedFaceIsBold()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
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
        var resolver = new DocxTests.SingleResolutionFontResolver(font.Value.Resolution);

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(2, DocxTests.CountPdfTextShows(pdf));
    }

    public static void DocxRendererDoesNotSynthesizeBoldForImplicitDefaultBoldFace()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
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
                      <w:r>
                        <w:rPr><w:b/></w:rPr>
                        <w:t>Default bold face</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var resolver = new DocxTests.SingleResolutionFontResolver(font.Value.Resolution);

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(2, DocxTests.CountPdfTextShows(pdf));
    }
    public static void DocxCffFontSubstitutesFallbackWithDiagnostic()
    {
        byte[] cffBytes = TestFontBuilder.CreateCffKindFont("CffFamily");
        byte[] fallbackBytes = TestFontBuilder.CreateTestFont();
        var resolver = new DocxTests.CffPrimaryFontResolver(cffBytes, fallbackBytes);
        var run = new DocxTextRun("Hello", 12d, null, false, false, false, null, "CffFamily")
        {
            Fonts = new DocxRunFonts("CffFamily", null, null, null, null, null, null, null)
        };
        DocxDocument document = DocxTests.CreateFontPlanDocument(run, new DocxFontCatalog([], DocxThemeFonts.Empty));
        var diagnostics = new List<OoxPdfDiagnostic>();
        IReadOnlyList<PdfPage> pages = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, diagnostics.Add, CancellationToken.None);

        TestAssert.Equal(1, pages.Count);
        PdfPage page = pages[0];
        TestAssert.True(DocxTests.CountPdfTextShows(page.Content) >= 1, "Substituted text must be painted.");
        TestAssert.Equal(1, page.Fonts.Count);
        TestAssert.True(page.Fonts[0].Font.Font.HasTrueTypeOutlines, "Emitted font must be embeddable.");
        TestAssert.Contains("TestFont", page.Fonts[0].Font.BaseFontName);
        OoxPdfDiagnostic warning = TestAssert.NotNull(diagnostics.SingleOrDefault(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"));
        TestAssert.DoesNotContain("CffFamily", page.Fonts[0].Font.BaseFontName);
        TestAssert.Equal(OoxPdfSeverity.Warning, warning.Severity);
        TestAssert.Contains("CffFamily", warning.Message);
        TestAssert.Equal("Document fallback typeface", warning.Fallback);
    }

    public static void DocxCffSubstitutionMatchesTrueTypeRendering()
    {
        byte[] cffBytes = TestFontBuilder.CreateCffKindFont("CffFamily");
        byte[] fallbackBytes = TestFontBuilder.CreateTestFont();
        var run = new DocxTextRun("Hello", 12d, null, false, false, false, null, "CffFamily")
        {
            Fonts = new DocxRunFonts("CffFamily", null, null, null, null, null, null, null)
        };
        DocxDocument document = DocxTests.CreateFontPlanDocument(run, new DocxFontCatalog([], DocxThemeFonts.Empty));
        var cffDiagnostics = new List<OoxPdfDiagnostic>();
        PdfPage cffPage = new DocxRenderer(new DocxTests.CffPrimaryFontResolver(cffBytes, fallbackBytes), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, cffDiagnostics.Add, CancellationToken.None)
            .Single();
        var ttDiagnostics = new List<OoxPdfDiagnostic>();
        PdfPage ttPage = new DocxRenderer(new DocxTests.CffPrimaryFontResolver(cffBytes, fallbackBytes, emitCff: false), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, ttDiagnostics.Add, CancellationToken.None)
            .Single();

        TestAssert.Equal(ttPage.Content, cffPage.Content);
        TestAssert.True(!ttDiagnostics.Any(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"), "TrueType rendering must not report a substitution.");
        TestAssert.True(cffDiagnostics.Any(diagnostic => diagnostic.Id == "FONT_UNSUPPORTED_OUTLINES"), "CFF rendering must report the substitution.");
    }

}
