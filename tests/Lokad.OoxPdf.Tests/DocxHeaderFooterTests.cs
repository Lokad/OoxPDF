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

internal static class DocxHeaderFooterTests
{
    public static void DocxLayoutStageOwnsSelectedStaticHeaderFooterLines()
    {
        DocxParagraph header = new(
            [
                new DocxTextRun("H", 10d, "FF0000", false, false, false, null, "Narrow"),
                new DocxTextRun("{PAGE}", 14d, "0000FF", false, false, false, null, "Narrow") { FieldKind = DocxFieldKind.Page }
            ],
            [],
            null,
            DocxTextAlignment.Center,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph footer = new(
            [new DocxTextRun("F", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Right,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = new(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            12d,
            12d,
            "240",
            "240",
            null,
            null,
            null,
            null)
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            },
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [footer]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxTextLineLayout[] staticLines = layout.Pages[0].StaticTextLines.ToArray();
        TestAssert.Equal(2, staticLines.Length);
        TestAssert.Equal("H1", staticLines[0].Text);
        TestAssert.Equal(95d, staticLines[0].X);
        TestAssert.Equal(174.84d, staticLines[0].BaselineY);
        TestAssert.Equal(2, staticLines[0].Segments.Count);
        TestAssert.Equal(10d, staticLines[0].Segments[0].FontSize ?? 0d);
        TestAssert.Equal(14d, staticLines[0].Segments[1].FontSize ?? 0d);
        TestAssert.Equal("1", staticLines[0].Segments[1].Text);
        TestAssert.Equal("0000FF", staticLines[0].Segments[1].StyleRun.ColorHex ?? string.Empty);
        TestAssert.Equal("F", staticLines[1].Text);
        TestAssert.Equal(185d, staticLines[1].X);
        TestAssert.Equal(12.6d, Math.Round(staticLines[1].BaselineY, 4));
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(2, snapshot.Pages[0].StaticTextLineCount);
        TestAssert.Equal(2, snapshot.Pages[0].StaticItems.Count);
        TestAssert.Equal("StaticHeaderTextLine", snapshot.Pages[0].StaticItems[0].Kind);
        TestAssert.Equal("StaticFooterTextLine", snapshot.Pages[0].StaticItems[1].Kind);
        TestAssert.Equal(2, snapshot.Pages[0].StaticStories.Count);
        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single(story => story.Kind == "Header");
        DocxStaticStoryLayoutSnapshot footerStory = snapshot.Pages[0].StaticStories.Single(story => story.Kind == "Footer");
        TestAssert.True(headerStory.VariantType == "default" && headerStory.TextLineCount == 1 && headerStory.ParagraphCount == 1 && headerStory.TextLength == 2 && headerStory.Items.Single().Kind == "StaticHeaderTextLine", "Static header story snapshots should group private-safe selected header line geometry.");
        TestAssert.True(footerStory.VariantType == "default" && footerStory.TextLineCount == 1 && footerStory.ParagraphCount == 1 && footerStory.TextLength == 1 && footerStory.Items.Single().Kind == "StaticFooterTextLine", "Static footer story snapshots should group private-safe selected footer line geometry.");
        TestAssert.Equal(1, snapshot.Pages[0].TextLineCount);
    }

    public static void DocxLayoutStageAddsStaticHeaderAfterEndnoteContinuationPages()
    {
        DocxParagraph header = new(
            [
                new DocxTextRun("H", 10d, "000000", false, false, false, null, "Narrow"),
                new DocxTextRun("{PAGE}", 10d, "000000", false, false, false, null, "Narrow") { FieldKind = DocxFieldKind.Page }
            ],
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
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "5",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote continuation line", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "5",
            Enumerable.Range(0, 8).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(header)]
            }
        };
        DocxDocument document = new(
            200d,
            100d,
            10d,
            10d,
            10d,
            10d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayoutSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);

        TestAssert.True(snapshot.Pages.Count >= 2, "The oversized endnote story should create a continuation page.");
        DocxLayoutPageSnapshot[] endnotePages = snapshot.Pages
            .Where(page => page.PlacedEndnoteStoryCount == 1)
            .ToArray();
        TestAssert.True(endnotePages.Length >= 1, "At least one page should own an endnote story slice.");
        foreach (DocxLayoutPageSnapshot endnotePage in endnotePages)
        {
            TestAssert.Equal(1, endnotePage.StaticTextLineCount);
            DocxLayoutItemSnapshot headerItem = endnotePage.StaticItems.Single(item => item.Kind == "StaticHeaderTextLine");
            TestAssert.Equal(2, headerItem.TextLength);
        }
    }

    public static void DocxLayoutStageSelectsStaticHeaderBodyElements()
    {
        DocxParagraph header = DocxTests.CreateDocxLayoutParagraph("HB", 10d, 10d);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(header)]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));

        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single();
        TestAssert.True(headerStory.Kind == "Header" && headerStory.VariantType == "default" && headerStory.TextLineCount == 1 && headerStory.ParagraphCount == 1 && headerStory.TextLength == 2, "Static header layout should select body elements directly instead of requiring the legacy paragraph map.");
        TestAssert.Equal("StaticHeaderTextLine", snapshot.Pages[0].StaticItems.Single().Kind);
    }

    public static void DocxLayoutStageResolvesStaticHeaderFloatingTextBoxNumPagesFieldFromPageCount()
    {
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph("Total {NUMPAGES}", 10d, 10d);
        DocxFloatingDrawing headerDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxParagraph firstBody = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph secondBody = DocxTests.CreateDocxLayoutParagraph("Second", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [headerDrawing]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(firstBody), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(secondBody)],
            [firstBody, secondBody],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        DocxFloatingDrawingLayout[] staticDrawings = layout.StaticFloatingDrawings.OrderBy(drawing => drawing.AnchorPageIndex).ToArray();
        TestAssert.Equal(2, staticDrawings.Length);
        TestAssert.Equal("Total 2", staticDrawings[0].TextBoxLayout!.TextLines.Single().Text);
        TestAssert.Equal("Total 2", staticDrawings[1].TextBoxLayout!.TextLines.Single().Text);
    }

    public static void DocxRendererTreatsStaticHeaderBodyElementsAsRenderableContent()
    {
        DocxParagraph headerCellParagraph = DocxTests.CreateDocxLayoutParagraph("HT", 10d, 10d);
        DocxTableCell headerCell = new("HT", [headerCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable headerTable = new(null, [40d], [new DocxTableRow([headerCell], 18d)]);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(headerTable)]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [],
            [],
            [headerTable]);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(DocxTests.CountPdfTextShows(page.Content) >= 1, "Static header body elements should pass the renderer's empty-document guard instead of producing a blank page.");
    }

    public static void DocxRendererTreatsStaticHeaderFloatingImagesAsRenderableContent()
    {
        var headerImage = new DocxInlineImage(72d, 36d, "image/png", TestFixtures.CreateRgbPng(1, 1, [32, 64, 96]), "/word/media/header.png");
        var headerDrawing = new DocxFloatingDrawing(
            DistanceTopValue: "0",
            DistanceBottomValue: "0",
            DistanceLeftValue: "0",
            DistanceRightValue: "0",
            SimplePositionValue: "0",
            RelativeHeightValue: "0",
            BehindDocumentValue: "0",
            LockedValue: null,
            LayoutInCellValue: null,
            AllowOverlapValue: null,
            ExtentCxValue: "914400",
            ExtentCyValue: "457200",
            HorizontalRelativeFromValue: "page",
            HorizontalAlignValue: null,
            HorizontalOffsetValue: "914400",
            VerticalRelativeFromValue: "page",
            VerticalAlignValue: null,
            VerticalOffsetValue: "457200",
            WrapKind: DocxFloatingWrapKind.None,
            WrapTextValue: null,
            ImageRelationshipId: "rIdHeaderImage1",
            Image: headerImage, SourceParagraphIndex: null, SourceBlockIndex: null);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [headerDrawing]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [],
            [],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(1, page.Images.Count);
        TestAssert.Contains("/Im1 Do", page.Content);
    }

    public static void DocxLayoutStageSummarizesSelectedStaticHeaderFooterVariants()
    {
        DocxParagraph defaultHeader = DocxTests.CreateDocxLayoutParagraph("DH", 10d, 10d);
        DocxParagraph firstHeader = DocxTests.CreateDocxLayoutParagraph("FH", 10d, 10d);
        DocxParagraph evenFooter = DocxTests.CreateDocxLayoutParagraph("EF", 10d, 10d);
        DocxParagraph firstBody = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph secondBody = DocxTests.CreateDocxLayoutParagraph("Second", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            TitlePage = true,
            EvenAndOddHeaders = true,
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
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(firstBody), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(secondBody)],
            [firstBody, secondBody],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));

        TestAssert.Equal(2, snapshot.Pages.Count);
        TestAssert.Equal("first", snapshot.Pages[0].StaticStories.Single().VariantType ?? string.Empty);
        TestAssert.True(!snapshot.Pages[1].StaticStories.Any(story => story.Kind == "Header"), "Even pages without an even header story must not fall back to the default header (w53 even-pages Office probe: Word renders those pages headerless).");
        TestAssert.Equal("even", snapshot.Pages[1].StaticStories.Single(story => story.Kind == "Footer").VariantType ?? string.Empty);
        TestAssert.True(snapshot.Pages[1].StaticItems.All(item => item.StoryVariantType is "default" or "even"), "Static item snapshots should retain the selected header/footer variant type.");
    }

    public static void DocxLayoutStageOwnsStaticHeaderInlineImages()
    {
        DocxInlineImage headerImage = new(36d, 18d, "image/png", [1, 2, 3], "/word/media/header.png");
        DocxParagraph header = new(
            [],
            [headerImage],
            null,
            DocxTextAlignment.Center,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxInlineImageLayout staticImage = layout.Pages[0].StaticInlineImages.Single();
        TestAssert.True(staticImage.Image == headerImage && staticImage.Width == 36d && staticImage.Height == 18d, "Static header inline images should be first-class page layout items with image geometry.");
        TestAssert.True(staticImage.SourceParagraphIndex == 0 && staticImage.StoryKind == "Header" && staticImage.StoryVariantType == "default", "Static header inline image layout should retain selected-story provenance.");

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(1, snapshot.Pages[0].StaticInlineImageCount);
        DocxLayoutItemSnapshot staticItem = snapshot.Pages[0].StaticItems.Single();
        TestAssert.True(staticItem.Kind == "StaticHeaderInlineImage" && staticItem.StoryVariantType == "default" && staticItem.SourceParagraphIndex == 0, "Static image snapshots should expose private-safe header story ownership.");
        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single();
        TestAssert.True(headerStory.Kind == "Header" && headerStory.TextLineCount == 0 && headerStory.InlineImageCount == 1 && headerStory.ParagraphCount == 1, "Static story snapshots should summarize inline image ownership separately from text line counts.");
    }

    public static void DocxLayoutStageWrapsStaticHeaderLines()
    {
        DocxParagraph header = new(
            [new DocxTextRun("Alpha Beta", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Center,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = new(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            12d,
            12d,
            "240",
            "240",
            null,
            null,
            null,
            null)
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            }
        };
        DocxDocument document = new(
            100d,
            200d,
            10d,
            50d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxTextLineLayout[] staticLines = layout.Pages[0].StaticTextLines.ToArray();
        TestAssert.Equal(2, staticLines.Length);
        TestAssert.Equal("Alpha ", staticLines[0].Text);
        TestAssert.Equal(15d, staticLines[0].X);
        TestAssert.Equal(178.6d, staticLines[0].BaselineY);
        TestAssert.Equal(0, staticLines[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, staticLines[0].SourceLineIndex ?? -1);
        TestAssert.True(staticLines[0].IsFirstParagraphLine == true, "The first wrapped static line should carry first-line ownership.");
        TestAssert.Equal("Beta", staticLines[1].Text);
        TestAssert.Equal(20d, staticLines[1].X);
        TestAssert.Equal(168.6d, staticLines[1].BaselineY);
        TestAssert.Equal(0, staticLines[1].SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, staticLines[1].SourceLineIndex ?? -1);
        TestAssert.True(staticLines[1].IsFirstParagraphLine == false, "Continuation static lines should not look like first paragraph lines.");

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(2, snapshot.Pages[0].StaticTextLineCount);
        TestAssert.Equal(2, snapshot.Pages[0].StaticItems.Count);
        TestAssert.Equal("StaticHeaderTextLine", snapshot.Pages[0].StaticItems[0].Kind);
        TestAssert.Equal(0, snapshot.Pages[0].StaticItems[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, snapshot.Pages[0].StaticItems[0].SourceLineIndex ?? -1);
        TestAssert.True(snapshot.Pages[0].StaticItems[0].IsFirstParagraphLine == true, "The static snapshot should preserve first-line ownership.");
        TestAssert.Equal(1, snapshot.Pages[0].StaticItems[1].SourceLineIndex ?? -1);
        TestAssert.True(snapshot.Pages[0].StaticItems[1].IsFirstParagraphLine == false, "The static snapshot should preserve continuation-line ownership.");
        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single();
        TestAssert.True(headerStory.Kind == "Header" && headerStory.TextLineCount == 2 && headerStory.ParagraphCount == 1 && headerStory.SourceLineCount == 2 && headerStory.FirstParagraphLineCount == 1 && headerStory.Items.Count == 2, "Static story snapshots should summarize wrapped selected header line ownership without exposing text.");
        TestAssert.Equal(0, headerStory.FirstSourceLineIndex ?? -1);
        TestAssert.Equal(1, headerStory.LastSourceLineIndex ?? -1);
    }

    public static void DocxLayoutStageAppliesStaticHeaderParagraphSpacing()
    {
        DocxParagraph first = new(
            [new DocxTextRun("A", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            6d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph second = new(
            [new DocxTextRun("B", 10d, null, false, false, false, null, "Narrow")],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            4d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = new(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            12d,
            12d,
            "240",
            "240",
            null,
            null,
            null,
            null)
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [first, second]
            }
        };
        DocxDocument document = new(
            100d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            []);

        DocxTextLineLayout[] staticLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .StaticTextLines
            .ToArray();

        TestAssert.Equal(2, staticLines.Length);
        TestAssert.Equal("A", staticLines[0].Text);
        TestAssert.Equal(178.6d, staticLines[0].BaselineY);
        TestAssert.Equal("B", staticLines[1].Text);
        TestAssert.Equal(162.6d, staticLines[1].BaselineY);
    }

    public static void DocxSyntheticHeaderAndFooterRenderOnPage()
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
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:footerReference w:type="default" r:id="rIdFooter1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Equal(6, DocxTests.CountPdfTextShows(pdf));
    }

    public static void DocxSyntheticHeaderReferenceTypesSelectDefaultWhenNoFirstOrEvenSetting()
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
                  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/header3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
                  <Relationship Id="rIdHeaderEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rIdHeaderDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                  <Relationship Id="rIdHeaderFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header3.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Even header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/header2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Default header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/header3.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="even" r:id="rIdHeaderEven"/>
                      <w:headerReference w:type="default" r:id="rIdHeaderDefault"/>
                      <w:headerReference w:type="first" r:id="rIdHeaderFirst"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal(3, document.HeaderParagraphsByType.Count);
            TestAssert.Equal(1, document.HeaderParagraphs.Count);
        }

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(4, DocxTests.CountPdfTextShows(pdf));
    }

    public static void DocxReaderPreservesHeaderFloatingDrawingsByVariant()
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
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """),
            ["word/_rels/header1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdHeaderImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/header.png"/>
                </Relationships>
                """),
            ["word/header1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:p>
                    <w:r>
                      <w:drawing>
                        <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="1">
                          <wp:extent cx="914400" cy="457200"/>
                          <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                          <wp:positionV relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                          <wp:wrapNone/>
                          <a:graphic>
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                              <pic:pic><pic:blipFill><a:blip r:embed="rIdHeaderImage1"/></pic:blipFill></pic:pic>
                            </a:graphicData>
                          </a:graphic>
                        </wp:anchor>
                      </w:drawing>
                    </w:r>
                  </w:p>
                </w:hdr>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/header.png"] = TestFixtures.CreateRgbPng(1, 1, [32, 64, 96])
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxFloatingDrawing drawing = document.HeaderFloatingDrawingsByType["default"].Single();
        TestAssert.Equal(1, document.PageSettings.HeaderFloatingDrawingsByType["default"].Count);
        TestAssert.Equal("rIdHeaderImage1", drawing.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/header.png", drawing.Image?.PartName ?? string.Empty);
        TestAssert.Equal(0, drawing.SourceParagraphIndex ?? -1);
        TestAssert.True(drawing.SourceBlockIndex is null, "Header floating drawings should not pretend to belong to a body block.");

        DocxStructureStorySnapshot story = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document).Stories.Single(story => story.Kind == "Header" && story.VariantType == "default");
        TestAssert.True(story.FloatingDrawingCount == 1 && story.ParagraphCount == 1, "Static header story snapshots should expose anchored drawing ownership without rendering it yet.");

        DocxLayoutSnapshot layout = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);
        DocxFloatingDrawingLayoutSnapshot layoutDrawing = layout.StaticFloatingDrawings.Single();
        TestAssert.True(layout.FloatingDrawings.Count == 0 && layoutDrawing.StoryKind == "Header" && layoutDrawing.StoryVariantType == "default", "Selected header drawings should be laid out in the static drawing stream, not mixed with body floating drawings.");
        TestAssert.Equal(0, layoutDrawing.AnchorPageIndex ?? -1);
        TestAssert.Equal(72d, layoutDrawing.PlacedX ?? 0d);
        TestAssert.Equal(756d, layoutDrawing.PlacedTop ?? 0d);
        TestAssert.Equal("Offset", layoutDrawing.HorizontalPlacementSource ?? string.Empty);
        TestAssert.Equal("Offset", layoutDrawing.VerticalPlacementSource ?? string.Empty);
        TestAssert.Equal(72d, layoutDrawing.ExtentWidthPoints ?? 0d);
        TestAssert.Equal(36d, layoutDrawing.ExtentHeightPoints ?? 0d);

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void DocxSyntheticHeaderFooterDistancesUsePageMarginTokens()
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
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header distance</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer distance</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:footerReference w:type="default" r:id="rIdFooter1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="1440" w:footer="1080"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] leftAlignedBaselines = Regex.Matches(pdf, @"1 0 0 1 36(?:\.\d+)? (?<y>-?\d+(?:\.\d+)?) Tm")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .ToArray();
        TestAssert.True(leftAlignedBaselines.Any(y => y > 704d && y < 720d), "Header baseline should be inset from the raw header-distance top by resolved font ascender metrics.");
        TestAssert.True(leftAlignedBaselines.Any(y => y > 63d && y < 70d), "Footer baseline should bottom-anchor the trailing cursor at the raw footer distance (R1: distance plus line-remainder plus after-spacing).");
        TestAssert.DoesNotContain("1 0 0 1 36 720 Tm", pdf);
        TestAssert.DoesNotContain("1 0 0 1 36 54 Tm", pdf);
    }

    public static void DocxSyntheticSectionHeadersUsePageLocalGeometry()
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
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                  <Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>Section header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>Section footer</w:t></w:r></w:p></w:ftr>
                """,
            ["word/header2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>Final header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="FF00FF"/></w:rPr><w:t>Final footer</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>First body</w:t></w:r></w:p>
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:headerReference w:type="default" r:id="rIdHeader1"/>
                          <w:footerReference w:type="default" r:id="rIdFooter1"/>
                          <w:pgSz w:w="4000" w:h="4000"/>
                          <w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360" w:header="360" w:footer="360"/>
                          <w:type w:val="nextPage"/>
                        </w:sectPr>
                      </w:pPr>
                    </w:p>
                    <w:p><w:r><w:t>Second body</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader2"/>
                      <w:footerReference w:type="default" r:id="rIdFooter2"/>
                      <w:pgSz w:w="6000" w:h="6000"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="1080" w:footer="1080"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] firstSectionLeftBaselines = DocxTests.ExtractTextBaselinesAtX(pdf, 18d);
        double[] finalSectionLeftBaselines = DocxTests.ExtractTextBaselinesAtX(pdf, 72d);
        TestAssert.True(firstSectionLeftBaselines.Any(y => y > 168d && y < 182d), "First-section header should use the first section header distance and left margin.");
        TestAssert.True(firstSectionLeftBaselines.Any(y => y > 27d && y < 34d), "First-section footer should bottom-anchor its trailing cursor at the first section footer distance.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 232d && y < 246d), "Final-section header should use the final section header distance and left margin.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 63d && y < 70d), "Final-section footer should bottom-anchor its trailing cursor at the final section footer distance.");
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
        TestAssert.Contains("1 0 1 rg", pdf);
    }

    public static void DocxSyntheticSectionHeadersDoNotBackfillEarlierSectionsFromFinalReferences()
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
                  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdHeader2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
                </Relationships>
                """,
            ["word/header2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="00FF00"/></w:rPr><w:t>Final header</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:color w:val="FF00FF"/></w:rPr><w:t>Final footer</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>First body</w:t></w:r></w:p>
                    <w:p>
                      <w:pPr>
                        <w:sectPr>
                          <w:pgSz w:w="4000" w:h="4000"/>
                          <w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360" w:header="360" w:footer="360"/>
                          <w:type w:val="nextPage"/>
                        </w:sectPr>
                      </w:pPr>
                    </w:p>
                    <w:p><w:r><w:t>Second body</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader2"/>
                      <w:footerReference w:type="default" r:id="rIdFooter2"/>
                      <w:pgSz w:w="6000" w:h="6000"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="1080" w:footer="1080"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        double[] firstSectionLeftBaselines = DocxTests.ExtractTextBaselinesAtX(pdf, 18d);
        double[] finalSectionLeftBaselines = DocxTests.ExtractTextBaselinesAtX(pdf, 72d);
        TestAssert.True(!firstSectionLeftBaselines.Any(y => y > 168d && y < 182d), "Final-section header should not backfill the earlier section that omits a header reference.");
        TestAssert.True(!firstSectionLeftBaselines.Any(y => y > 18d && y < 30d), "Final-section footer should not backfill the earlier section that omits a footer reference.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 232d && y < 246d), "Final-section header should still render on its owning section.");
        TestAssert.True(finalSectionLeftBaselines.Any(y => y > 63d && y < 70d), "Final-section footer should still render on its owning section.");
    }

    public static void DocxEmptyHeaderDisplacesNothing()
    {
        // Empty header stories expose positive infinity so the renderer two-pass keeps
        // the single-pass layout untouched (table-fragment regression 2026-09-08: an
        // empty header cursor inside the header zone read as overflow whenever the
        // header distance was smaller than the top margin).
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(double.PositiveInfinity, layout.HeaderContentBottomByPage[0]);
    }

    public static void DocxOverflowingFooterTopIsExposedOnLayout()
    {
        // Supports footer-overflow frame raising: two tokened footer paras advance
        // 11.5833 each inside an R1 bottom-anchored block (first baseline 33.7667),
        // so the content top (first baseline plus the 9.4 body-rule offset) lands at
        // 43.1667 on the layout for the renderer two-pass to raise the frame by.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Fa"), StaticPara("Fb")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(43.1667, Math.Round(layout.FooterContentTopByPage[0], 4));
    }

    public static void DocxBodyBreaksBeforeRaisedFooterFrameBottom()
    {
        // Office A/B (w48 multi-line footer plus long body, Word-COM rendered): body
        // breaks before footer content instead of overlapping it (30 body lines on page
        // one, then the break). The footer frame map raises the page bottom, so the
        // second 10pt body para no longer fits under the first on a 100pt page.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph first = StaticPara("One");
        DocxParagraph second = StaticPara("Two");
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 10d,
            FooterDistancePoints = 10d,
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("F")]
            }
        };
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(first), new DocxParagraphElement(second)], [first, second], []);
        var footerMap = new Dictionary<int, double> { [0] = 60d };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None, null, null, footerMap);
        TestAssert.Equal(2, layout.Pages.Count);
        DocxTextLineLayout secondPageFirstLine = layout.Pages[1].Items.OfType<DocxTextLineLayout>().First();
        TestAssert.Equal(80.6, Math.Round(secondPageFirstLine.BaselineY, 4));
    }

    public static void DocxBodyTrailingAfterDoesNotMoveFooter()
    {
        // Office A/B (w49 body-trailing probe, Word-COM rendered): body trailing
        // after-spacing is swallowed at the body-to-footer boundary (footer at 47.90
        // with body after-24 and after-0 alike). Guards against future cross-boundary
        // tuck regressions: the footer baseline must not depend on body trailing.
        static DocxParagraph BodyPara(double afterPoints, string? afterValue) => new(
            [new DocxTextRun("Body", 12d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            afterPoints,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, afterValue, null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        static DocxParagraph FootPara() => new(
            [new DocxTextRun("Foot", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        static double FootBaseline(double bodyAfterPoints, string? bodyAfterValue)
        {
            DocxPageSettings settings = DocxPageSettings.Empty with
            {
                HeaderDistancePoints = 20d,
                FooterDistancePoints = 20d,
                FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = [FootPara()]
                }
            };
            DocxParagraph body = BodyPara(bodyAfterPoints, bodyAfterValue);
            DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            return layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Footer").Single().BaselineY;
        }

        TestAssert.Equal(Math.Round(FootBaseline(0d, "0"), 4), Math.Round(FootBaseline(24d, "480"), 4));
    }

    public static void DocxFooterLeadingBeforeDoesNotMoveFooter()
    {
        // Office A/B (w50 footer-before probe, Word-COM rendered): footer-first
        // before-spacing is absorbed by the R1 bottom-anchor shift (footer at 47.90
        // with before-24 and before-0 alike). Guards the anchor against regressions.
        static DocxParagraph FootPara(double beforePoints, string? beforeValue) => new(
            [new DocxTextRun("Foot", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            beforePoints,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(beforeValue, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        static double FootBaseline(double beforePoints, string? beforeValue)
        {
            DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
            DocxPageSettings settings = DocxPageSettings.Empty with
            {
                HeaderDistancePoints = 20d,
                FooterDistancePoints = 20d,
                FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = [FootPara(beforePoints, beforeValue)]
                }
            };
            DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            return layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Footer").Single().BaselineY;
        }

        TestAssert.Equal(Math.Round(FootBaseline(0d, null), 4), Math.Round(FootBaseline(24d, "480"), 4));
    }

    public static void DocxReaderReadsEvenAndOddHeadersFromSectionProperties()
    {
        // Office A/B (w53 even-pages probe, Word-COM rendered): Word honors the
        // section-level w:evenAndOddHeaders flag (page-two stories suppressed without
        // even variants), but the reader only consulted settings.xml, so section-level
        // flags never reached story selection.
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
                    <w:p><w:r><w:t>Body</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:evenAndOddHeaders/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        TestAssert.True(document.FinalSectionBreak?.PageSettings.EvenAndOddHeaders == true, "Section-level evenAndOddHeaders should resolve.");
    }

    public static void DocxTitlePageSuppressesDefaultHeaderWithoutFirstStory()
    {
        // Office A/B (w52 title-page probe, Word-COM rendered): with titlePg set and no
        // first story defined, page one shows no header (Word does not fall back to the
        // default story), while page-one footers behave the same way. Guards the story
        // selection against default-fallback regressions.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            TitlePage = true,
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("H")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(0, layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Header").Count());
    }

    public static void DocxEvenPagesSuppressDefaultHeaderWithoutEvenStory()
    {
        // Office A/B (w53 even-pages probe, Word-COM rendered): with even/odd headers on
        // and no even story defined, even pages show no header (Word does not fall back
        // to the default story there either), while odd pages keep the default story.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = new(
            [new DocxTextRun(string.Concat(Enumerable.Repeat("word ", 80)), 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            EvenAndOddHeaders = true,
            HeaderDistancePoints = 10d,
            FooterDistancePoints = 10d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("H")]
            }
        };
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.True(layout.Pages.Count > 1, "The long body should paginate.");
        TestAssert.Equal(1, layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Header").Count());
        TestAssert.Equal(0, layout.Pages[1].StaticTextLines.Where(line => line.StoryKind == "Header").Count());
    }

    public static void DocxOverflowingHeaderBottomIsExposedOnLayout()
    {
        // Supports header-overflow displacement: three tokened 10pt header paras
        // advance 11.5833 each from startY 180, so the trailing content bottom lands
        // at 145.25 on the layout for the renderer two-pass to displace by.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Ha"), StaticPara("Hb"), StaticPara("Hc")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(145.25, Math.Round(layout.HeaderContentBottomByPage[0], 4));
    }

    public static void DocxBodyStartsBelowOverflowingHeaderContent()
    {
        // Office A/B (w37/w39 multi-paragraph headers, Word-COM rendered): body starts
        // below overflowing header content (w39 body top equals header content bottom
        // 681.39). The displacement map pins the page start cursor, so the 10pt body
        // first baseline lands at 135.85 instead of the undisplaced 180.6.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = new(
            [new DocxTextRun("Body", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            278d / 240d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Ha"), StaticPara("Hb"), StaticPara("Hc")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        var displacement = new Dictionary<int, double> { [0] = 34.75d };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None, null, displacement);
        DocxTextLineLayout bodyLine = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single();
        TestAssert.Equal(135.85, Math.Round(bodyLine.BaselineY, 4));
    }

    public static void DocxBodyAppliesOwnBeforeBelowHeaderContentBottom()
    {
        // Office A/B (w47 cross-story spacing probe, Word-COM rendered): the header
        // trailing after-spacing does not cross the story boundary (header after-24
        // plus body before-12 applies 12, landing 682.06). The displaced page start is
        // the header trailing cursor including trailing after-spacing; the body first
        // paragraph then applies its own before-spacing fresh.
        static DocxParagraph HeaderPara(string text, double afterPoints, string? afterValue) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            afterPoints,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, afterValue, null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = new(
            [new DocxTextRun("Body", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            12d,
            8d,
            278d / 240d,
            null,
            new DocxParagraphSpacing("240", null, null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [HeaderPara("Ha", 0d, "0"), HeaderPara("Hb", 24d, "480")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout undisplaced = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        double headerBottom = undisplaced.HeaderContentBottomByPage[0];
        TestAssert.Equal(132.8333, Math.Round(headerBottom, 4));
        var displacement = new Dictionary<int, double> { [0] = 180d - 132.8333d };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None, null, displacement);
        DocxTextLineLayout bodyLine = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single();
        TestAssert.Equal(111.4333, Math.Round(bodyLine.BaselineY, 4));
    }

    public static void DocxRecordedZeroDisplacementBeatsFallback()
    {
        // Office A/B (w53 even-pages probe, Word-COM rendered): page two carries no
        // stories, so it starts at the top with a full frame even though page one
        // displaces. Explicit zero entries must win over the beyond-map last-entry
        // fallback (which only covers pages past the first layout); a sparse map that
        // drops the zero would wrongly displace page two by the page-one entry.
        DocxParagraph body = new(
            [new DocxTextRun(string.Concat(Enumerable.Repeat("word ", 80)), 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxPageSettings settings = DocxPageSettings.Empty;
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        var headerMap = new Dictionary<int, double> { [0] = 11.5833d, [1] = 0d };
        var footerMap = new Dictionary<int, double> { [0] = 60d, [1] = 0d };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None, null, headerMap, footerMap);
        TestAssert.True(layout.Pages.Count > 1, "The long body should paginate.");
        DocxTextLineLayout secondPageFirstLine = layout.Pages[1].Items.OfType<DocxTextLineLayout>().First();
        TestAssert.Equal(80.6, Math.Round(secondPageFirstLine.BaselineY, 4));
    }

    public static void DocxBodyDisplacementFallsBackToLastEntryBeyondMap()
    {
        // Pages beyond a short displacement map reuse its last entry (extra pages arise
        // from the displacement itself under a repeated header), so page two of a long
        // body starts displaced by the page-zero entry here.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = new(
            [new DocxTextRun(string.Concat(Enumerable.Repeat("word ", 80)), 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 10d,
            FooterDistancePoints = 10d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("H")]
            }
        };
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        var displacement = new Dictionary<int, double> { [0] = 11.5833d };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None, null, displacement);
        TestAssert.True(layout.Pages.Count > 1, "The long body should paginate.");
        DocxTextLineLayout secondPageFirstLine = layout.Pages[1].Items.OfType<DocxTextLineLayout>().First();
        TestAssert.Equal(69.0167, Math.Round(secondPageFirstLine.BaselineY, 4));
    }

    public static void DocxStaticHeaderAutoAdvanceFollowsSingleLineFactor()
    {
        // Office A/B (w37-staticfree Final-mode probe, Word-COM rendered): static auto
        // lineHeight is single-height times 278/240 (Arial10 pitch 13.32 vs 13.321,
        // Aptos16 22.56 vs 22.624), while the layout used windows extents (12.0 here);
        // header origin follows the body baseline rule (w39 bare-11 inset 10.44 vs
        // body-bare 10.46), not winAscent (which sits 0.31 high at this size).
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Ha"), StaticPara("Hb")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] headerLines = layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Header").ToArray();
        TestAssert.Equal(2, headerLines.Length);
        TestAssert.Equal(170.6, Math.Round(headerLines[0].BaselineY, 4));
        TestAssert.Equal(11.5833, Math.Round(headerLines[0].BaselineY - headerLines[1].BaselineY, 4));
    }

    public static void DocxStaticFooterAutoAdvanceBottomAnchorsBlock()
    {
        // Office A/B (w38 default, w39 tokened and w40 bare multi-line footers at
        // distances 36/72/54 plus w43-solo single-line bare footer, Word-COM rendered):
        // footer auto advances follow the single-times-factor law and the text-only
        // block bottom-anchors its trailing cursor at the footer distance, so this
        // two-line tokened block lands its first baseline at 33.7667 with the 11.5833
        // pitch preserved.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Fa"), StaticPara("Fb")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] footerLines = layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Footer").ToArray();
        TestAssert.Equal(2, footerLines.Length);
        TestAssert.Equal(33.7667, Math.Round(footerLines[0].BaselineY, 4));
        TestAssert.Equal(11.5833, Math.Round(footerLines[0].BaselineY - footerLines[1].BaselineY, 4));
    }

    public static void DocxStaticFooterFollowerLinesUseBodyRuleOffsets()
    {
        // Office A/B (w38/w39/w40 footer probes, Word-COM rendered): footer lines use
        // body-rule offsets inside a bottom-anchored block (Ften-b to Fsix-a gaps read
        // 19.08 to 19.11 for a 18.96 body-rule prediction). FamilyWidth arithmetic
        // here: top-down first baseline 10.6, advances 11.5833 and 18.5333, then the
        // block shift of 30.1167 pins the trailing cursor at the footer distance 20.
        static DocxParagraph StaticPara(string text, double fontSize) => new(
            [new DocxTextRun(text, fontSize, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Fa", 10d), StaticPara("Fb", 16d)]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] footerLines = layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Footer").ToArray();
        TestAssert.Equal(2, footerLines.Length);
        TestAssert.Equal(40.7167, Math.Round(footerLines[0].BaselineY, 4));
        TestAssert.Equal(23.4933, Math.Round(footerLines[1].BaselineY, 4));
    }

    public static void DocxStaticHeaderAutoAdvanceScalesUniformlyWithSpacing()
    {
        // Office A/B (w36 WC re-baselines, Word-COM balloon refs): Word scales static
        // advances uniformly with the lane scale (single-x-factor times s, after-steps
        // agree to 0.03), while the layout scaled line metrics by the 0.7936 fitted
        // compromise (9.5232 here) and spacing by s. The production-plumbing mirror
        // (scaled layout measurer plus raw fallback) must advance by raw single times
        // factor times the spacing scale.
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            278d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, null, null, null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Ha"), StaticPara("Hb")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        var raw = new DocxTests.FamilyWidthTextMeasurer();
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(raw, 0.76d, 0.79359971328d);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 0.76d).Create(document, scaled, CancellationToken.None, raw);
        DocxTextLineLayout[] headerLines = layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Header").ToArray();
        TestAssert.Equal(2, headerLines.Length);
        TestAssert.Equal(170.6, Math.Round(headerLines[0].BaselineY, 4));
        TestAssert.Equal(8.8033, Math.Round(headerLines[0].BaselineY - headerLines[1].BaselineY, 4));
    }

    public static void DocxStaticHeaderExplicitLineFactorFollowsSingleLineFactor()
    {
        // Office A/B (w44 explicit-factor probe, Word-COM rendered): explicit w:line
        // factors scale the single height too (Arial12 E100 pitch 13.80 vs 13.80, E115
        // pitch 15.86 vs 15.871), and explicit-rule origin follows the body baseline
        // rule (locked explicit-100 header2 inset 9.36 vs 9.40). Only exact/atLeast
        // statics keep legacy behavior (unprobed).
        static DocxParagraph StaticPara(string text) => new(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            276d / 240d,
            null,
            new DocxParagraphSpacing(null, "0", null, null, null, null, "276", "auto", null),
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [StaticPara("Ha"), StaticPara("Hb")]
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)], [body], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] headerLines = layout.Pages[0].StaticTextLines.Where(line => line.StoryKind == "Header").ToArray();
        TestAssert.Equal(2, headerLines.Length);
        TestAssert.Equal(170.6d, Math.Round(headerLines[0].BaselineY, 4));
        TestAssert.Equal(11.5d, Math.Round(headerLines[0].BaselineY - headerLines[1].BaselineY, 4));
    }
}
