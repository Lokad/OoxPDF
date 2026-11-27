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

internal static partial class DocxHeaderFooterTests
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
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 12d,
            FooterDistancePoints = 12d,
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(header)
            },
            FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(footer)
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
            [new DocxParagraphElement(body)]);

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

    public static void DocxStaticHeaderMixedSizePitchTransitionsScaleWithPrintScale()
    {
        // RV06 pagination probe (edge-mixedstory-wc, Word 16.0): under a
        // word-compatible print scale, header 12->15 exceeds by 0.75 and 15->12
        // undershoots by 0.69. Pre-fix the scaled layout matches preserve.
        DocxParagraph SizedLine(string text, double fontSize)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, fontSize, 12d) with
            {
                LineSpacingPoints = null
            };
        }

        double[] Pitches(OoxPdfDocxMarkupGeometryMode mode, double scale)
        {
            DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
            DocxPageSettings settings = DocxPageSettings.Empty with
            {
                HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = DocxTests.CreateParagraphBlocks(SizedLine("Alpha", 12d), SizedLine("Beta", 15d), SizedLine("Gamma", 12d))
                }
            };
            DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
            DocxTextLineLayout[] lines = new DocxLayoutEngine(mode, scale)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.Single().StaticTextLines.ToArray();
            TestAssert.Equal(3, lines.Length);
            return [lines[0].BaselineY - lines[1].BaselineY, lines[1].BaselineY - lines[2].BaselineY];
        }

        // Static line heights scale through the spacing scale even in unit tests,
        // so the assertion is whole-pitch scaling (transition included).
        double[] preserve = Pitches(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, 1d);
        double[] scaled = Pitches(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 0.75d);
        TestAssert.True(
            Math.Abs(scaled[0] - preserve[0] * 0.75d) < 0.001d,
            $"Header 12->15 pitch should scale uniformly. preserve={preserve[0]} scaled={scaled[0]}.");
        TestAssert.True(
            Math.Abs(scaled[1] - preserve[1] * 0.75d) < 0.001d,
            $"Header 15->12 pitch should scale uniformly. preserve={preserve[1]} scaled={scaled[1]}.");
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
            Enumerable.Range(0, 8).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(), null);
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
            [new DocxParagraphElement(body)])
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
            [new DocxParagraphElement(body)]);

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
            [new DocxParagraphElement(firstBody), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(secondBody)]);

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
        DocxTableCell headerCell = new([headerCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
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
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();

        TestAssert.True(DocxTests.CountPdfTextShows(page.Content) >= 1, "Static header body elements should pass the renderer's empty-document guard instead of producing a blank page.");
    }

    public static void DocxRendererTreatsStaticHeaderFloatingImagesAsRenderableContent()
    {
        const string uri = "https://example.invalid/header-picture";
        var headerImage = new DocxInlineImage(72d, 36d, "image/png", TestFixtures.CreateRgbPng(1, 1, [32, 64, 96]), "/word/media/header.png")
        {
            ClickTarget = uri
        };
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
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();

        TestAssert.Equal(1, page.Images.Count);
        TestAssert.Contains("/Im1 Do", page.Content);
        var link = page.Annotations.Single();
        TestAssert.Equal(uri, link.Uri);
        TestAssert.Equal((72d, 128d, 72d, 36d), (link.X, link.Y, link.Width, link.Height));
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
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(defaultHeader),
                ["first"] = DocxTests.CreateParagraphBlocks(firstHeader)
            },
            FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["even"] = DocxTests.CreateParagraphBlocks(evenFooter)
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
            [new DocxParagraphElement(firstBody), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(secondBody)]);

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
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(header)
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
            [new DocxParagraphElement(body)]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxInlineImageLayout staticImage = layout.Pages[0].StaticInlineImages.Single();
        TestAssert.True(staticImage.Image == headerImage && staticImage.Width == 36d && staticImage.Height == 18d, "Static header inline images should be first-class page layout items with image geometry.");
        TestAssert.True(staticImage.SourceParagraphIndex == 0 && staticImage.Story?.Kind == DocxStoryKind.Header && staticImage.Story?.VariantType == "default", "Static header inline image layout should retain selected-story provenance.");

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
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 12d,
            FooterDistancePoints = 12d,
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(header)
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
            [new DocxParagraphElement(body)]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxTextLineLayout[] staticLines = layout.Pages[0].StaticTextLines.ToArray();
        TestAssert.Equal(2, staticLines.Length);
        TestAssert.Equal("Alpha ", staticLines[0].Text);
        // RV06 wrapalign probe: Office centers drawable text even on wrapped
        // non-last lines (break space excluded from the centering width).
        TestAssert.Equal(17.5d, staticLines[0].X);
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
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 12d,
            FooterDistancePoints = 12d,
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(first, second)
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
            [new DocxParagraphElement(body)]);

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
        var lines = ReadStaticStoryLines("text");
        TestAssert.Equal(3, lines.Length);
        TestAssert.Equal("Header text", lines[0].Text);
        TestAssert.Equal("Body text", lines[1].Text);
        TestAssert.Equal("Footer text", lines[2].Text);
        TestAssert.True(lines.All(line => line.X == 72d && line.Y > 0d && line.Y < 792d),
            "The three stories should start at the default left margin inside the authored page.");
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
            TestAssert.Equal(3, document.HeaderBodyElementsByType.Count);
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
                  <Relationship Id="rIdEarlierHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
                  <Relationship Id="rIdMissingHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="missing-header.xml"/>
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
                    <w:r><w:t>Selected header</w:t></w:r>
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
            ["word/header2.xml"] = TestFixtures.Utf8("""
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:p><w:r><w:t>Earlier header</w:t></w:r></w:p>
                </w:hdr>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdEarlierHeader"/>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:headerReference w:type="default" r:id="rIdMissingHeader"/>
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

        var headerBlocks = document.HeaderBodyElementsByType["default"];
        TestAssert.True(ReferenceEquals(headerBlocks, document.PageSettings.HeaderBodyElementsByType["default"]),
            "The document and section references share the parsed header blocks.");
        var header = DocxBlockTraversal.EnumerateDirectParagraphs(headerBlocks).Single();
        TestAssert.Equal("Selected header", string.Concat(header.Runs.Select(run => run.Text)));
        DocxFloatingDrawing drawing = document.HeaderFloatingDrawingsByType["default"].Single();
        TestAssert.Equal(1, document.PageSettings.HeaderFloatingDrawingsByType["default"].Count);
        TestAssert.Equal("rIdHeaderImage1", drawing.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/header.png", drawing.Image?.PartName ?? string.Empty);
        TestAssert.Equal(0, drawing.SourceParagraphIndex ?? -1);
        TestAssert.True(drawing.SourceBlockIndex is null, "Header floating drawings should not pretend to belong to a body block.");

        DocxStructureStorySnapshot story = DocxStructureSnapshot.FromDocument(document).Stories.Single(story => story.Kind == "Header" && story.VariantType == "default");
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
        var lines = ReadStaticStoryLines("distance",
            """<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="1440" w:footer="1080"/>""");
        TestAssert.Equal(3, lines.Length);
        TestAssert.Equal("Header distance", lines[0].Text);
        TestAssert.Equal("Body text", lines[1].Text);
        TestAssert.Equal("Footer distance", lines[2].Text);
        TestAssert.True(lines.All(line => line.X == 36d), "All three stories should start at the authored left margin.");
        TestAssert.True(lines[0].Y is > 704d and < 720d,
            "The header baseline should be inset from the authored header distance.");
        TestAssert.True(lines[2].Y is > 63d and < 70d,
            "The footer baseline should include the trailing line remainder and paragraph spacing.");
    }

    private static (string Text, double X, double Y)[] ReadStaticStoryLines(string suffix, string pageMargins = "")
    {
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
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
            ["_rels/.rels"] = TestFixtures.DocxPackageRelationships,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header {suffix}</w:t></w:r></w:p></w:hdr>
                """,
            ["word/footer1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer {suffix}</w:t></w:r></w:p></w:ftr>
                """,
            ["word/document.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:footerReference w:type="default" r:id="rIdFooter1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      {pageMargins}
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using var package = OoxPackage.Open(input, default);
        const OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.Final;
        DocxDocument document = new DocxReader().Read(package, null, default, mode);
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        PdfPage page = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
        return ReadStaticStoryLines(page);
    }

    private static (string Text, double X, double Y)[] ReadStaticStoryLines(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, honorActualText: false)
            .GroupBy(show => show.Y)
            .OrderByDescending(line => line.Key)
            .Select(line => (Text: string.Concat(line.Select(show => show.Text)).Trim(), X: line.Min(show => show.X), Y: line.Key))
            .Where(line => line.Text.Length > 0)
            .ToArray();

    public static void DocxSyntheticSectionHeadersUsePageLocalGeometry()
    {
        PdfPage[] pages = RenderSyntheticSectionStories(firstSectionHasStories: true);
        var lines = ReadStaticStoryLines(pages[0]);
        AssertSectionStory(lines, "Section header", 18d, 168d, 182d);
        AssertSectionStory(lines, "Section footer", 18d, 27d, 34d);
        TestAssert.Contains("1 0 0 rg", pages[0].Content);
        TestAssert.Contains("0 0 1 rg", pages[0].Content);
        AssertFinalSectionStories(pages[1]);
    }

    public static void DocxSyntheticSectionHeadersDoNotBackfillEarlierSectionsFromFinalReferences()
    {
        PdfPage[] pages = RenderSyntheticSectionStories(firstSectionHasStories: false);
        TestAssert.Equal("First body", ReadStaticStoryLines(pages[0]).Single().Text);
        AssertFinalSectionStories(pages[1]);
    }

    private static void AssertFinalSectionStories(PdfPage page)
    {
        var lines = ReadStaticStoryLines(page);
        AssertSectionStory(lines, "Final header", 72d, 232d, 246d);
        AssertSectionStory(lines, "Final footer", 72d, 63d, 70d);
        TestAssert.Contains("0 1 0 rg", page.Content);
        TestAssert.Contains("1 0 1 rg", page.Content);
    }

    private static void AssertSectionStory((string Text, double X, double Y)[] lines,
        string text, double x, double lowerBaseline, double upperBaseline)
    {
        var story = lines.Single(line => line.Text == text);
        TestAssert.Equal(x, story.X);
        TestAssert.True(story.Y > lowerBaseline && story.Y < upperBaseline,
            $"The painted {text} baseline {story.Y} belongs inside its authored section frame.");
    }

    private static PdfPage[] RenderSyntheticSectionStories(bool firstSectionHasStories)
    {
        string firstReferences = firstSectionHasStories ? StoryReferences(1) : string.Empty;
        var parts = new Dictionary<string, string>
        {
            ["_rels/.rels"] = TestFixtures.DocxPackageRelationships,
            ["word/document.xml"] = $$"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>First body</w:t></w:r></w:p>
                    <w:p><w:pPr><w:sectPr>
                      {{firstReferences}}
                      <w:pgSz w:w="4000" w:h="4000"/>
                      <w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360" w:header="360" w:footer="360"/>
                      <w:type w:val="nextPage"/>
                    </w:sectPr></w:pPr></w:p>
                    <w:p><w:r><w:t>Second body</w:t></w:r></w:p>
                    <w:sectPr>
                      {{StoryReferences(2)}}
                      <w:pgSz w:w="6000" w:h="6000"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="1080" w:footer="1080"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        };
        if (firstSectionHasStories)
        {
            parts.Add("word/header1.xml", StoryXml("hdr", "Section header", "FF0000"));
            parts.Add("word/footer1.xml", StoryXml("ftr", "Section footer", "0000FF"));
        }
        parts.Add("word/header2.xml", StoryXml("hdr", "Final header", "00FF00"));
        parts.Add("word/footer2.xml", StoryXml("ftr", "Final footer", "FF00FF"));

        var stories = new[] { 1, 2 }.Where(index => index == 2 || firstSectionHasStories)
            .SelectMany(index => new[] { (Kind: "header", Index: index), (Kind: "footer", Index: index) }).ToArray();
        string overrides = string.Concat(stories.Select(story =>
            $"""<Override PartName="/word/{story.Kind}{story.Index}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{story.Kind}+xml"/>"""));
        parts.Add("[Content_Types].xml", $$"""
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              {{overrides}}
            </Types>
            """);
        string relationships = string.Concat(stories.Select(story =>
            $"""<Relationship Id="{story.Kind}{story.Index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{story.Kind}" Target="{story.Kind}{story.Index}.xml"/>"""));
        parts.Add("word/_rels/document.xml.rels", $$"""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              {{relationships}}
            </Relationships>
            """);

        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        const OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.Final;
        var document = new DocxReader().Read(package, null, default, mode);
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        PdfPage[] pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(2, pages.Length);
        return pages;

        static string StoryReferences(int index) => $$"""
            <w:headerReference w:type="default" r:id="header{{index}}"/>
            <w:footerReference w:type="default" r:id="footer{{index}}"/>
            """;
        static string StoryXml(string root, string text, string color) => $$"""
            <w:{{root}} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:p><w:r><w:rPr><w:color w:val="{{color}}"/></w:rPr><w:t>{{text}}</w:t></w:r></w:p>
            </w:{{root}}>
            """;
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
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(double.PositiveInfinity, layout.HeaderContentBottomByPage[0]);
    }

    private static DocxParagraph CreateStaticParagraph(string text) => new(
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

    public static void DocxOverflowingFooterTopIsExposedOnLayout()
    {
        // Supports footer-overflow frame raising: two tokened footer paras advance
        // 11.5833 each inside an R1 bottom-anchored block (first baseline 33.7667),
        // so the content top (first baseline plus the 9.4 body-rule offset) lands at
        // 43.1667 on the layout for the renderer two-pass to raise the frame by.
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("Fa"), CreateStaticParagraph("Fb"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(43.1667, Math.Round(layout.FooterContentTopByPage[0], 4));
    }

    public static void DocxBodyBreaksBeforeRaisedFooterFrameBottom()
    {
        // Office A/B (w48 multi-line footer plus long body, Word-COM rendered): body
        // breaks before footer content instead of overlapping it (30 body lines on page
        // one, then the break). The footer frame map raises the page bottom, so the
        // second 10pt body para no longer fits under the first on a 100pt page.
        DocxParagraph first = CreateStaticParagraph("One");
        DocxParagraph second = CreateStaticParagraph("Two");
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 10d,
            FooterDistancePoints = 10d,
            FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("F"))
            }
        };
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(first), new DocxParagraphElement(second)]);
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
                FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = DocxTests.CreateParagraphBlocks(FootPara())
                }
            };
            DocxParagraph body = BodyPara(bodyAfterPoints, bodyAfterValue);
            DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            return layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Footer).Single().BaselineY;
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
                FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = DocxTests.CreateParagraphBlocks(FootPara(beforePoints, beforeValue))
                }
            };
            DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            return layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Footer).Single().BaselineY;
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
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            TitlePage = true,
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("H"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(0, layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Header).Count());
    }

    public static void DocxEvenPagesSuppressDefaultHeaderWithoutEvenStory()
    {
        // Office A/B (w53 even-pages probe, Word-COM rendered): with even/odd headers on
        // and no even story defined, even pages show no header (Word does not fall back
        // to the default story there either), while odd pages keep the default story.
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
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("H"))
            }
        };
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.True(layout.Pages.Count > 1, "The long body should paginate.");
        TestAssert.Equal(1, layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Header).Count());
        TestAssert.Equal(0, layout.Pages[1].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Header).Count());
    }

    public static void DocxOverflowingHeaderBottomIsExposedOnLayout()
    {
        // Supports header-overflow displacement: three tokened 10pt header paras
        // advance 11.5833 each from startY 180, so the trailing content bottom lands
        // at 145.25 on the layout for the renderer two-pass to displace by.
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("Ha"), CreateStaticParagraph("Hb"), CreateStaticParagraph("Hc"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(145.25, Math.Round(layout.HeaderContentBottomByPage[0], 4));
    }

    public static void DocxBodyStartsBelowOverflowingHeaderContent()
    {
        // Office A/B (w37/w39 multi-paragraph headers, Word-COM rendered): body starts
        // below overflowing header content (w39 body top equals header content bottom
        // 681.39). The displacement map pins the page start cursor, so the 10pt body
        // first baseline lands at 135.85 instead of the undisplaced 180.6.
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
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("Ha"), CreateStaticParagraph("Hb"), CreateStaticParagraph("Hc"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
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
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(HeaderPara("Ha", 0d, "0"), HeaderPara("Hb", 24d, "480"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
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
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(body)]);
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
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("H"))
            }
        };
        DocxDocument document = new(100d, 100d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(body)]);
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
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("Ha"), CreateStaticParagraph("Hb"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] headerLines = layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Header).ToArray();
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
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("Fa"), CreateStaticParagraph("Fb"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] footerLines = layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Footer).ToArray();
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
            FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(StaticPara("Fa", 10d), StaticPara("Fb", 16d))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] footerLines = layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Footer).ToArray();
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
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 20d,
            FooterDistancePoints = 20d,
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(CreateStaticParagraph("Ha"), CreateStaticParagraph("Hb"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        var raw = new DocxTests.FamilyWidthTextMeasurer();
        var scaled = new DocxRenderer.ScaledDocxTextMeasurer(raw, 0.76d, 0.79359971328d);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 0.76d).Create(document, scaled, CancellationToken.None, raw);
        DocxTextLineLayout[] headerLines = layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Header).ToArray();
        TestAssert.Equal(2, headerLines.Length);
        // The natural frame exposes printed coordinates: source baseline170.6
        // maps about the200pt page center at scale0.76. Its pitch is unchanged.
        TestAssert.True(headerLines.All(line => line.StaticStoryUsesPhysicalY), "The complete natural story owns its printed origin.");
        TestAssert.Equal(153.656, Math.Round(headerLines[0].BaselineY, 4));
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
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(StaticPara("Ha"), StaticPara("Hb"))
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 20d, 20d, settings, [], [], [], [new DocxParagraphElement(body)]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] headerLines = layout.Pages[0].StaticTextLines.Where(line => line.Story?.Kind == DocxStoryKind.Header).ToArray();
        TestAssert.Equal(2, headerLines.Length);
        TestAssert.Equal(170.6d, Math.Round(headerLines[0].BaselineY, 4));
        TestAssert.Equal(11.5d, Math.Round(headerLines[0].BaselineY - headerLines[1].BaselineY, 4));
    }

    public static void DocxHeaderTrailingLinesShareDrawableLinePositions()
    {
        // RV06 header-align probe: Office centers/rights drawable header text and
        // keeps one row-end space beyond authored trailing in headers too.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">H trail   </w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">H trail</w:t></w:r></w:p></w:hdr>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxTextLineLayout[] headerLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .StaticTextLines
            .ToArray();
        DocxTextLineLayout trailLine = headerLines.Single(line => line.Text.EndsWith("   ", StringComparison.Ordinal));
        DocxTextLineLayout cleanLine = headerLines.Single(line => line.Text == "H trail");
        TestAssert.Equal("H trail    ", trailLine.Text);
        TestAssert.Equal(cleanLine.X, trailLine.X);
    }

    public static void DocxStaticHeaderFirstBaselineUsesHheaAscenderWhenLarger()
    {
        // Word 16.0 header grid (Calibri/Tahoma 10pt first baselines 746.52 vs
        // 746.04): static first baselines join the hhea-ascender opt-in (Tahoma gap
        // 0.56 matches hheaAsc 1.0005 over Calibri 0.9521). The patched synthetic
        // face pins hheaAsc at 1.1em, so the header first baseline must sit 1.6pt
        // below its legacy position.
        DocxTextRun run = new("H", 10d, null, false, false, false, null, null);
        DocxParagraph header = new([run], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderDistancePoints = 18d, // This control measures hhea metrics, independently of source-default distance.
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = DocxTests.CreateParagraphBlocks(header)
            }
        };
        DocxDocument document = new(200d, 200d, 10d, 10d, 10d, 10d, settings, [], [], [], [new DocxParagraphElement(body)]);
        var resolver = new DocxTests.TallHheaFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("HheaFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxTextLineLayout[] staticLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0].StaticTextLines.ToArray();
        TestAssert.Equal(1, staticLines.Length);
        TestAssert.True(Math.Abs(staticLines[0].BaselineY - 171.0d) < 0.000001d, "Header first baseline must follow max hhea ascender when larger; observed baseline=" + staticLines[0].BaselineY.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxStaticHeaderFirstBaselineFollowsMaxHheaRun()
    {
        // RV05 headermix probes (Word COM references edge-headermix-cal/tah): static
        // headers follow the same max-hhea rule as body text (a direct Tahoma 12pt run
        // shifts the header baseline by 0.73); widest-by-size layout keeps both equal.
        double calBaseline = LayoutStaticHeaderBaselineMixedRun("Calibri");
        double tahBaseline = LayoutStaticHeaderBaselineMixedRun("Tahoma");
        TestAssert.True(Math.Abs((calBaseline - tahBaseline) - 0.73d) < 0.05d, "Static header baseline must follow the max-hhea run; observed shift=" + (calBaseline - tahBaseline).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class MixedHheaTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsAscender(run, fontSize) : fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize)
        {
            double em = run?.FontFamily == "Tahoma" ? 1.0005d : 0.75d;
            return em * fontSize;
        }
    }

    private static double LayoutStaticHeaderBaselineMixedRun(string midFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Hdr start words </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="FAM" w:hAnsi="FAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">middle mixed words</w:t></w:r></w:p></w:hdr>""".Replace("FAM", midFamily),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new MixedHheaTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages.Single().StaticTextLines.Single().BaselineY;
    }

    }
