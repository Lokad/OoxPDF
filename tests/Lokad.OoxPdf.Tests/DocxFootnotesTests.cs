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

internal static class DocxFootnotesTests
{
    public static void DocxWordCompatibleAllMarkupClampsCommentConnectorAnchorsAtPageEdge()
    {
        DocxTextRun run = new("Edge anchor", 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        DocxParagraph paragraph = new(
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
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 1, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ],
            CommentRanges =
            [
                new DocxCommentRange("1", 0, 0, 0, 0, 1, 0)
            ]
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public edge comment body", 10d, 12d))],
            [],
            [], null);
        // Near-edge clamp calibration: anchor = L - L * (1 - s) - 3.18 with s = 200 / 346.5,
        // so L = 4 lands at -0.87 and clamps to 0.5 (L = 10 no longer reaches the edge
        // under the lane-fit print scale).
        DocxDocument document = new(
            200d,
            240d,
            4d,
            120d,
            20d,
            20d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxMarkupBalloonPlacementSnapshot placement = renderer.InspectMarkupBalloons(document)
            .Single(item => item.Kind == "Comment");

        TestAssert.True(
            placement.AnchorConnectorX >= 0.5d &&
            placement.AnchorConnectorX < 1d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Word-compatible all-markup should clamp near-edge comment connector anchors inside the page media box. AnchorX={placement.AnchorConnectorX}."));
        TestAssert.True(placement.AnchorConnectorClamped, "Near-edge connector placement snapshots should expose that the anchor was clamped.");
    }

    public static void DocxReaderSplitsNestedVisibleInlineContainerPageBreaks()
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/break" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:hyperlink r:id="rIdLink">
                        <w:ins>
                          <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                        </w:ins>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Text before the nested break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Nested page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Text after the nested break should remain a paragraph fragment.");
        DocxParagraph before = ((DocxParagraphElement)document.BodyElements[0]).Paragraph;
        DocxParagraph after = ((DocxParagraphElement)document.BodyElements[2]).Paragraph;
        TestAssert.Equal("Before", string.Concat(before.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(after.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, before.Hyperlinks.Count);
        TestAssert.Equal(1, after.Hyperlinks.Count);
    }

    public static void DocxReaderSplitsMoveToFinalViewPageBreaks()
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
                      <w:moveTo w:id="9" w:author="Author" w:date="2026-06-02T00:00:00Z">
                        <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                      </w:moveTo>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Moved-to text before the break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Moved-to page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Moved-to text after the break should remain a paragraph fragment.");
        TestAssert.Equal("Before", string.Concat(((DocxParagraphElement)document.BodyElements[0]).Paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(((DocxParagraphElement)document.BodyElements[2]).Paragraph.Runs.Select(run => run.Text)));
    }

    public static void DocxLayoutReservesPageFootnoteStoryArea()
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor paragraph", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "7",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler paragraph keeps body close to the bottom", 10d, 12d);
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body line one wraps line two", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "7",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(anchor),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler)
            ],
            [],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot footnotePage = snapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 1);
        double footnoteTop = footnotePage.PlacedRelatedItems.Max(item => item.Y + item.Height);
        double bodyBottom = footnotePage.Items.Min(item => item.Y);

        TestAssert.True(footnotePage.PlacedRelatedStoryTextLineCount >= 1, "Resolved footnote story text should be placed on the body-reference page.");
        TestAssert.True(bodyBottom >= footnoteTop, $"Body layout should reserve page space above placed footnotes; body bottom {bodyBottom} footnote top {footnoteTop}.");
    }

    public static void DocxFootnoteStoryLaysOutUnscaledDesignSpace()
    {
        // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote stories lay
        // out unscaled in design space while emission maps uniformly (WC first
        // baseline 236.62 = affine-mapped design). Pre-fix the word-compatible
        // story layout scaled metrics in place.
        DocxParagraph SizedLine(string text, double fontSize)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, fontSize, 12d) with
            {
                LineSpacingPoints = null
            };
        }

        double[] Pitches(OoxPdfDocxMarkupGeometryMode mode, double scale)
        {
            DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d) with
            {
                InlineReferences =
                [
                    new DocxInlineReference(
                        DocxRelatedStoryKind.Footnote,
                        "7",
                        CustomMarkFollowsValue: null,
                        DisplayText: "1",
                        SourceRunIndex: 0,
                        RunChildIndex: 1,
                        TextOffsetInRun: 6)
                ]
            };
            var footnoteStory = new DocxRelatedStory(
                DocxRelatedStoryKind.Footnote,
                "/word/footnotes.xml",
                "7",
                [
                    new DocxParagraphElement(SizedLine("Alpha", 12d)),
                    new DocxParagraphElement(SizedLine("Beta", 15d)),
                    new DocxParagraphElement(SizedLine("Gamma", 12d))
                ],
                [],
                [],
                null);
            var document = new DocxDocument(
                220d,
                200d,
                10d,
                10d,
                10d,
                10d,
                DocxPageSettings.Empty,
                [],
                [],
                [],
                [new DocxParagraphElement(anchor)],
                [],
                [])
            {
                RelatedStories = [footnoteStory]
            };
            DocxTextLineLayout[] lines = new DocxLayoutEngine(mode, scale)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.SelectMany(page => page.PlacedRelatedStories).SelectMany(story => story.TextLines).ToArray();
            TestAssert.Equal(3, lines.Length);
            return [lines[0].BaselineY - lines[1].BaselineY, lines[1].BaselineY - lines[2].BaselineY];
        }

        // Footnote stories lay out unscaled in design space (emission maps
        // uniformly); the scaled layout keeps full unscaled transitions.
        double[] preserve = Pitches(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, 1d);
        double[] scaled = Pitches(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 0.75d);
        TestAssert.True(
            Math.Abs(scaled[0] - preserve[0]) < 0.001d,
            $"Footnote story layout should stay design space. preserve={preserve[0]} scaled={scaled[0]}.");
        TestAssert.True(
            Math.Abs(scaled[1] - preserve[1]) < 0.001d,
            $"Footnote story layout should stay design space. preserve={preserve[1]} scaled={scaled[1]}.");
    }

    public static void DocxEndnoteStartTopRecoversDesignBodyBottom()
    {
        // RV06 anchor probe (edge-endanchor-5, Word 16.0): scaled body bottom
        // 689.77 with frame top 720, first inset 11.28 and print scale 0.75874
        // recovers the design body end 683.75; scale 1.0 and degenerate frames
        // keep legacy behavior.
        TestAssert.True(
            Math.Abs(DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(689.77d, 720d, 11.28d, 0.75874d) - 683.75d) < 0.02d,
            "Design body end should recover 683.75.");
        TestAssert.Equal(100d, DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(100d, 720d, 11.28d, 1d));
        TestAssert.Equal(700d, DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(700d, 720d, null, 0.75d));
    }

    public static void DocxLayoutStacksMultipleFootnotesOnOnePage()
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor paragraph with two footnote markers", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "31",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 8),
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "32",
                    CustomMarkFollowsValue: null,
                    DisplayText: "2",
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 22)
            ]
        };
        DocxParagraph firstFootnote = DocxTests.CreateDocxLayoutParagraph("First footnote body", 10d, 12d);
        DocxParagraph secondFootnote = DocxTests.CreateDocxLayoutParagraph("Second footnote body", 10d, 12d);
        var firstStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "31",
            [new DocxParagraphElement(firstFootnote)],
            [],
            [], null);
        var secondStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "32",
            [new DocxParagraphElement(secondFootnote)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(anchor)], []) with
        {
            RelatedStories = [firstStory, secondStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot footnotePage = snapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 2);
        DocxPlacedRelatedStoryLayoutSnapshot[] placedStories = footnotePage.PlacedRelatedStories.ToArray();
        DocxPlacedRelatedStoryLayoutSnapshot firstPlaced = placedStories.Single(story => story.Id == "31");
        DocxPlacedRelatedStoryLayoutSnapshot secondPlaced = placedStories.Single(story => story.Id == "32");

        TestAssert.True(firstPlaced.TopY > secondPlaced.TopY, "The first footnote body should be placed above the later footnote body instead of both stories sharing the page bottom.");
        TestAssert.True(secondPlaced.TopY <= firstPlaced.TopY - firstPlaced.Height + 0.001d, "Stacked footnote bodies should not overlap vertically.");
        TestAssert.True(firstPlaced.SeparatorY is not null && secondPlaced.SeparatorY is null, "Generic separator geometry should be emitted once above the stacked footnote group.");
    }

    public static void DocxRendererClipsOverlongPlacedFootnoteStoryToPageRegion()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("body with long footnote", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "16",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote overflow body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "16",
            Enumerable.Range(0, 8).Select(_ => new DocxParagraphElement(footnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [bodyParagraph],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxPlacedRelatedStoryLayoutSnapshot placedStory = snapshot.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .First(story => story.Kind == "Footnote" && story.Id == "16");
        // Overlong notes slice across pages now; the head slice keeps full-story height with a page-slice height.
        PdfPage placedPage = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, null, CancellationToken.None)
            .First(page => page.Content.Contains(" re W n", StringComparison.Ordinal));

        TestAssert.True(placedStory.ContentHeight > placedStory.Height, "The overlong footnote should retain full-story height while exposing the clipped page slice height.");
        TestAssert.Equal(0d, placedStory.ContentTopOffset);
        TestAssert.Contains(" re W n", placedPage.Content);
    }

    public static void DocxLayoutPlacesFootnoteOnRenderedMarkerRunPage()
    {
        var leadingRun = new DocxTextRun(
            string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12)),
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        var markerRun = new DocxTextRun(
            "marker",
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        DocxParagraph anchor = new(
            [leadingRun, markerRun],
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
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "11",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 1,
                    RunChildIndex: 1,
                    TextOffsetInRun: 0)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "11",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The marker paragraph should span pages so note ownership is not reducible to source-block ownership.");
        TestAssert.True(layout.Pages[0].Items.OfType<DocxTextLineLayout>().Any(line => line.SourceBlockIndex == 0), "The first page should contain the source paragraph before the marker run is rendered.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page where the marker run is rendered, not the first page containing the source paragraph.");
        TestAssert.True(layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line => line.Segments.Any(segment => segment.SourceTextRunIndex == 1)), "The selected footnote page should contain the marker source run.");
    }

    public static void DocxLayoutPlacesFootnoteOnRenderedMarkerOffsetPage()
    {
        const int markerOffset = 180;
        string text = string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12));
        var run = new DocxTextRun(
            text,
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        DocxParagraph anchor = new(
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
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "12",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: markerOffset)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "12",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The single marker run should span pages so note ownership requires source offsets.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page containing the marker offset, not the first page containing the source run.");
        TestAssert.True(
            layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line =>
                line.Segments.Any(segment =>
                    segment.SourceTextRunIndex == 0 &&
                    markerOffset >= segment.SourceTextOffsetInRun &&
                    markerOffset <= segment.SourceTextOffsetInRun + segment.Text.Length)),
            "The selected footnote page should contain the marker source offset inside the rendered run segment.");
    }

    public static void DocxLayoutPlacesFootnoteOnReaderProvenanceMarkerOffsetPage()
    {
        string leadingText = string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12));
        const string markerText = "1";
        int markerOffset = leadingText.Length;
        var leadingRun = new DocxTextRun(
            leadingText,
            10d,
            null,
            false,
            false,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        var markerRun = new DocxTextRun(
            markerText,
            10d,
            null,
            false,
            false,
            false,
            null,
            null,
            0d,
            false,
            "superscript",
            false,
            null,
            false,
            null,
            null,
            null,
            null,
            null,
            false,
            null,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = markerOffset
        };
        var trailingRun = new DocxTextRun(
            " trailing text after marker",
            10d,
            null,
            false,
            false,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = markerOffset
        };
        DocxParagraph anchor = new(
            [leadingRun, markerRun, trailingRun],
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
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "13",
                    CustomMarkFollowsValue: null,
                    DisplayText: markerText,
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: markerOffset)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "13",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The normalized run sequence should span pages so source provenance is required.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page containing the normalized marker run, not the first page containing earlier text from the same source run.");
        TestAssert.True(
            layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line =>
                line.Segments.Any(segment =>
                    segment.Text == markerText &&
                    segment.SourceTextRunIndex == 0 &&
                    segment.SourceTextOffsetInRun == markerOffset)),
            "The selected footnote page should contain the normalized marker display run with original OOXML source provenance.");
    }

    public static void DocxLayoutPlacesSectEndEndnoteOnOwningSectionEndPage()
    {
        DocxParagraph firstSectionParagraph = DocxTests.CreateDocxLayoutParagraph("first section endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "21",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph secondSectionParagraph = DocxTests.CreateDocxLayoutParagraph("second section body", 10d, 12d);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote body", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "21",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        var sectionBreak = new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []);
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(firstSectionParagraph),
                sectionBreak,
                new DocxParagraphElement(secondSectionParagraph)
            ],
            [firstSectionParagraph, secondSectionParagraph],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot endnotePage = snapshot.Pages.Single(page => page.PlacedEndnoteStoryCount == 1);

        TestAssert.True(snapshot.Pages.Count >= 2, "The next-page section break should separate the two sections.");
        TestAssert.Equal("sectEnd", snapshot.Pages[0].SectionEndnotePositionValue ?? string.Empty);
        TestAssert.Equal(1, snapshot.Pages[0].PlacedEndnoteStoryCount);
        TestAssert.Equal(0, snapshot.Pages[^1].PlacedEndnoteStoryCount);
        TestAssert.True(endnotePage.PlacedRelatedStories.Any(story => story.Kind == "Endnote" && story.SourceBlockIndex == 0), "A sectEnd endnote should be owned by the source section end page, not rewritten as a document-end story.");
    }

    public static void DocxLayoutWrapsDocumentEndEndnoteUsingFinalSectionPageWidth()
    {
        DocxPageSettings wideFirstSection = new(
            "12240",
            "12000",
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
        DocxPageSettings narrowFinalSection = new(
            "4400",
            "6000",
            null,
            "200",
            "200",
            "200",
            "200",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Narrow document-end endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "25",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 7)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("alpha beta gamma delta epsilon zeta eta theta", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "25",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            narrowFinalSection,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide first section", 10d, 12d)),
                new DocxSectionBreakElement(wideFirstSection, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(anchor)
            ],
            [],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "25");
        DocxTextLineLayout[] lines = placedStory.TextLines.ToArray();

        TestAssert.Equal(-1, placedStory.SourceBlockIndex);
        TestAssert.True(Math.Abs(placedStory.Width - 200d) < 0.001d, "The document-end endnote should inherit the final section body width.");
        TestAssert.True(lines.Length >= 2, "Document-end endnotes should wrap using the target page width, not the first section width.");
        TestAssert.True(lines.All(line => line.Width <= placedStory.Width + 0.001d), "Document-end endnote lines should not exceed the final section body width.");
    }

    public static void DocxLayoutWrapsSectEndEndnoteUsingOwningSectionPageWidth()
    {
        DocxPageSettings wideFirstSection = new(
            "12240",
            "12000",
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
        DocxPageSettings narrowFinalSection = new(
            "4400",
            "6000",
            null,
            "200",
            "200",
            "200",
            "200",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null)
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Narrow section-end endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "26",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 7)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("alpha beta gamma delta epsilon zeta eta theta", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "26",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            narrowFinalSection,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide first section", 10d, 12d)),
                new DocxSectionBreakElement(wideFirstSection, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(anchor)
            ],
            [],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "26");
        DocxTextLineLayout[] lines = placedStory.TextLines.ToArray();

        TestAssert.Equal(2, placedStory.SourceBlockIndex);
        TestAssert.True(Math.Abs(placedStory.Width - 200d) < 0.001d, "The section-end endnote should inherit the owning section body width.");
        TestAssert.True(lines.Length >= 2, "Section-end endnotes should wrap using the target section width, not the first section width.");
        TestAssert.True(lines.All(line => line.Width <= placedStory.Width + 0.001d), "Section-end endnote lines should not exceed the owning section body width.");
    }

    public static void DocxLayoutKeepsSectEndEndnoteOverflowPageBeforeFollowingSection()
    {
        DocxParagraph firstSectionParagraph = DocxTests.CreateDocxLayoutParagraph("first section endnote marker", 10d, 18d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "22",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6),
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "23",
                    CustomMarkFollowsValue: null,
                    DisplayText: "2",
                    SourceRunIndex: 0,
                    RunChildIndex: 2,
                    TextOffsetInRun: 12)
            ]
        };
        DocxParagraph secondSectionParagraph = DocxTests.CreateDocxLayoutParagraph("second section body", 10d, 12d);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote overflow body", 10d, 18d);
        var firstEndnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "22",
            Enumerable.Range(0, 4).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        var secondEndnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "23",
            Enumerable.Range(0, 4).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(firstSectionParagraph),
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(secondSectionParagraph)
            ],
            [firstSectionParagraph, secondSectionParagraph],
            [])
        {
            RelatedStories = [firstEndnoteStory, secondEndnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        int secondSectionPageIndex = snapshot.Pages
            .Select((page, pageIndex) => (page, pageIndex))
            .First(item => item.page.Items.Any(pageItem => pageItem.SourceBlockIndex == 2))
            .pageIndex;
        int overflowEndnotePageIndex = snapshot.Pages
            .Select((page, pageIndex) => (page, pageIndex))
            .Where(item => item.page.PlacedEndnoteStoryCount == 1)
            .Select(item => item.pageIndex)
            .Last();

        TestAssert.True(snapshot.Pages.Count >= 3, "Multiple sectEnd endnotes that cannot all fit on the section end page should create a section-owned continuation page.");
        TestAssert.True(overflowEndnotePageIndex > 0 && overflowEndnotePageIndex < secondSectionPageIndex, "Section-end endnote overflow must be inserted before the following section instead of falling back to document-end placement.");
        TestAssert.Equal("sectEnd", snapshot.Pages[overflowEndnotePageIndex].SectionEndnotePositionValue ?? string.Empty);
        TestAssert.True(snapshot.Pages[overflowEndnotePageIndex].PlacedRelatedStories.Any(story => story.Kind == "Endnote" && story.SourceBlockIndex == 0), "The inserted section-end continuation page should retain marker-owned endnote provenance.");
        TestAssert.Equal(0, snapshot.Pages[secondSectionPageIndex].PlacedEndnoteStoryCount);
    }

    public static void RelatedStoryPageIndexMatchesLegacyMatching()
    {
        // R12: the once-per-pass page index answers reference/block queries exactly
        // like the legacy per-check page walks: null-tolerance and identity paragraph
        // matching, offset and run-fallback segment checks, vacuous truth for
        // never-rendered runs, and last-page section ranges.
        DocxParagraph source = MakeNoteParagraph();
        DocxParagraph other = MakeNoteParagraph();
        DocxTextLineLayout lineA = MakeOwnerLine(null, 0, 7, "hello");
        DocxTextLineLayout lineB = MakeOwnerLine(other, 0, 7, "world");
        DocxTextLineLayout lineC = MakeOwnerLine(null, 1, 7, "!");
        DocxLayoutPage page0 = MakeOwnerPage([lineA]);
        DocxLayoutPage page1 = MakeOwnerPage([lineB, lineC]);
        DocxLayoutEngine.RelatedStoryPageIndex index = DocxLayoutEngine.RelatedStoryPageIndex.Build([page0, page1], CancellationToken.None);

        DocxInlineReference reference = new(DocxRelatedStoryKind.Footnote, "1", null, null, SourceRunIndex: 7, RunChildIndex: 0, TextOffsetInRun: 2);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference);
        TestAssert.True(index.IsReferenceRenderedOnPage(0, location), "Null-tolerance line must match any paragraph of its block.");
        TestAssert.True(!index.IsReferenceRenderedOnPage(1, location), "Identity mismatch and block mismatch must reject page 1.");
        TestAssert.Equal(0, index.FindFirstPageWithReference(location));

        var block1 = new DocxLayoutEngine.DocxInlineReferenceLocation(1, source, reference with { TextOffsetInRun = 0 });
        TestAssert.True(!index.IsReferenceRenderedOnPage(0, block1), "Block mismatch must reject page 0.");
        TestAssert.True(index.IsReferenceRenderedOnPage(1, block1), "Null-tolerance line must match block 1 on page 1.");

        var missingBlock = new DocxLayoutEngine.DocxInlineReferenceLocation(9, source, reference);
        TestAssert.True(index.IsReferenceRenderedOnPage(0, missingBlock), "Blocks rendered nowhere stay vacuously true (legacy).");
        TestAssert.True(index.IsReferenceRenderedOnPage(1, missingBlock), "Blocks rendered nowhere stay vacuously true (legacy).");
        TestAssert.Equal(0, index.FindFirstPageWithReference(missingBlock));

        var unrenderedRun = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference with { SourceRunIndex = 99 });
        TestAssert.True(index.IsReferenceRenderedOnPage(1, unrenderedRun), "Never-rendered runs stay vacuously true (legacy).");
        var negativeRun = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference with { SourceRunIndex = -1 });
        TestAssert.True(index.IsReferenceRenderedOnPage(1, negativeRun), "Negative runs stay vacuously true (legacy).");

        int[] page0Blocks = index.SortedBlocks(0);
        TestAssert.Equal(1, page0Blocks.Length);
        TestAssert.Equal(0, page0Blocks[0]);
        int[] page1Blocks = index.SortedBlocks(1);
        TestAssert.Equal(2, page1Blocks.Length);
        TestAssert.Equal(0, page1Blocks[0]);
        TestAssert.Equal(1, page1Blocks[1]);
        TestAssert.Equal(1, index.FindLastPageWithBlockInRange(0, 0));
        TestAssert.Equal(1, index.FindLastPageWithBlockInRange(0, 1));
        TestAssert.Equal(-1, index.FindLastPageWithBlockInRange(5, 9));
    }

    private static DocxParagraph MakeNoteParagraph()
    {
        return new DocxParagraph(
            [],
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
            null);
    }

    private static DocxTextLineLayout MakeOwnerLine(DocxParagraph? paragraph, int block, int run, string text)
    {
        var styleRun = new DocxTextRun(text, 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout(
            text,
            styleRun,
            0d,
            text.Length * 5d,
            10d,
            0d,
            0d,
            default,
            false,
            run,
            0,
            DocxTextSegmentRole.Text);
        return new DocxTextLineLayout(
            Text: text,
            StyleRun: styleRun,
            FontSize: 10d,
            X: 0d,
            BaselineY: 10d,
            Width: text.Length * 5d,
            Segments: [segment],
            SourceBlockIndex: block,
            SourceParagraphIndex: null,
            SourceLineIndex: null,
            Story: null,
            LineHeight: null,
            AppliedBeforeSpacing: null,
            IsFirstParagraphLine: null,
            EndsWithIntraTokenBreak: false,
            SingleLineHeight: null,
            ListLabelSingleLineHeight: null,
            BodyWindowsLineHeight: null,
            ListLabelWindowsLineHeight: null,
            EffectiveLineSpacingFactor: null,
            LineSpacingFactorFloorApplied: null,
            PendingAfterSpacing: null,
            ParagraphBeforeSpacing: null,
            ParagraphAfterSpacing: null,
            ContextualSpacingSuppressed: null,
            SourceParagraph: paragraph,
            LineHeightSource: null,
            EmitsTerminalParagraphMark: false);
    }

    private static DocxLayoutPage MakeOwnerPage(IReadOnlyList<DocxTextLineLayout> lines)
    {
        return new DocxLayoutPage(
            612d,
            792d,
            72d,
            72d,
            0d,
            72d,
            72d,
            DocxPageSettings.Empty,
            new DocxSectionLayoutProperties(null, null, null, null, null, null, []),
            [],
            [],
            [],
            [],
            [],
            lines);
    }

    public static void DocxFootnoteTrailingLinesShareDrawableLinePositions()
    {
        // RV06 footnote-align probe: Office centers/rights drawable footnote text and
        // keeps one row-end space beyond authored trailing in footnotes too.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">F trail   </w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">F trail</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxTextLineLayout[] footnoteLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .SelectMany(story => story.TextLines)
            .ToArray();
        DocxTextLineLayout trailLine = footnoteLines.Single(line => line.Text.EndsWith("   ", StringComparison.Ordinal));
        DocxTextLineLayout cleanLine = footnoteLines.Single(line => line.Text == "F trail");
        TestAssert.Equal("F trail    ", trailLine.Text);
        TestAssert.Equal(cleanLine.X, trailLine.X);
    }

    public static void DocxFootnoteSeparatorMarkSitsAtRuleEnd()
    {
        // RV06 footnote probe: Office renders the separator mark as a space at the
        // rule end (body-left plus the 144pt rule width).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        TestAssert.Equal(1, separatorLine.Segments.Count);
        TestAssert.Equal(" ", separatorLine.Segments[0].Text);
        TestAssert.Equal(144d, separatorLine.Segments[0].X - separatorLine.X);
    }

    private sealed class StrikeoutFontResolver : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            byte[] faceBytes = TestFontBuilder.CreateTestFont();
            PatchStrikeoutMetrics(faceBytes, size: 50, position: 300);
            return new FontFaceResolution(
                request.FamilyName,
                "StrikeFace",
                new FontStyleKey(request.Bold, request.Italic),
                new MemoryFontProgramSource("test:strikeface", faceBytes),
                IsFallback: false);
        }

        private static void PatchStrikeoutMetrics(byte[] faceBytes, ushort size, short position)
        {
            // The shared synthetic face stores a placeholder strikeout size, so pin
            // real OS/2 strikeout geometry (0.05em size, 0.30em position) in place.
            int tableCount = (faceBytes[4] << 8) | faceBytes[5];
            for (int tableIndex = 0; tableIndex < tableCount; tableIndex++)
            {
                int record = 12 + 16 * tableIndex;
                if (faceBytes[record] == 0x4F && faceBytes[record + 1] == 0x53 && faceBytes[record + 2] == 0x2F && faceBytes[record + 3] == 0x32)
                {
                    int offset = (faceBytes[record + 8] << 24) | (faceBytes[record + 9] << 16) | (faceBytes[record + 10] << 8) | faceBytes[record + 11];
                    faceBytes[offset + 26] = (byte)(size >> 8);
                    faceBytes[offset + 27] = (byte)(size & 0xFF);
                    faceBytes[offset + 28] = (byte)((position >> 8) & 0xFF);
                    faceBytes[offset + 29] = (byte)(position & 0xFF);
                    return;
                }
            }

            throw new InvalidOperationException("Synthetic test font is missing the OS/2 table.");
        }
    }

    public static void DocxFootnoteSeparatorRuleFollowsStrikeoutMetrics()
    {
        // RV06 separator probes (Word 16.0, Times, Aptos and Calibri at 10, 12 and 14pt):
        // the footnote separator rule follows OS/2 strikeout geometry, with the rule top
        // at the strikeout position and the thickness at the strikeout size, instead of
                // The patched synthetic face pins strikeout 0.30em and 0.05em,
        // so the rule bottom must sit 0.25em above the separator bottom while the mark baseline
        // rides at the separator bottom with no extra ride.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double markSize = separatorLine.Segments[0].StyleRun.EffectiveProperties.FontSize;
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - 0.25d * markSize) < 0.000001d, "Footnote rule bottom must sit 0.25em above the mark baseline with no ride.");
        TestAssert.True(Math.Abs(separator.SeparatorThickness - 0.05d * markSize) < 0.000001d, "Footnote rule thickness must follow the strikeout size.");
    }

    public static void DocxFootnoteSeparatorRuleKeepsLegacyConstantsWithoutStrikeoutMetrics()
    {
        // Measurers without strikeout metrics keep the legacy footnote constants.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - 2.1d) < 0.000001d, "Legacy footnote rule offset must stay 2.1pt above the separator bottom with no ride.");
        TestAssert.Equal(0.75d, separator.SeparatorThickness);
    }

    public static void DocxEndnoteSeparatorRuleFollowsStrikeoutMetrics()
    {
        // RV06 endnote probes match the footnote strikeout rule, so section-end endnote
        // separators share the same font-derived geometry instead of a fixed 3.74pt offset.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double markSize = separatorLine.Segments[0].StyleRun.EffectiveProperties.FontSize;
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - (0.25d * markSize - 0.15d)) < 0.000001d, "Endnote rule bottom must sit 0.25em above the mark baseline.");
        TestAssert.True(Math.Abs(separator.SeparatorThickness - 0.05d * markSize) < 0.000001d, "Endnote rule thickness must follow the strikeout size.");
    }

    public static void DocxWordCompatibleAllMarkupMapsSeparatorRuleThickness()
    {
        // RV06 endnote probes: word-compatible separator rules map uniformly to
        // emission space like the surrounding story text, but the rule thickness
        // still emits at design size while the rule width is print-scaled.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:commentRangeStart w:id="0"/></w:r><w:r><w:t xml:space="preserve">commented</w:t></w:r><w:r><w:commentRangeEnd w:id="0"/></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:initials="R" w:date="2024-01-02T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });;
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.AllMarkup);
        }
        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        double designThickness = separator.SeparatorThickness;
        double designWidth = Math.Min(144d, separator.Width);
        var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
        Match[] ruleMatches = Regex.Matches(page.Content, @"(?<x>-?[0-9.]+) (?<y>-?[0-9.]+) (?<w>[0-9.]+) (?<h>[0-9.]+) re f[^*]")
            .Cast<Match>()
            .Where(match => double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture) < 200d && double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture) > 50d)
            .ToArray();
        TestAssert.Equal(1, ruleMatches.Length);
        double emitWidth = double.Parse(ruleMatches[0].Groups["w"].Value, CultureInfo.InvariantCulture);
        double emitHeight = double.Parse(ruleMatches[0].Groups["h"].Value, CultureInfo.InvariantCulture);
        double printScale = emitWidth / designWidth;
        TestAssert.True(printScale < 0.999d, "Separator rule width must carry the print scale, proving the emission map is active.");
        TestAssert.True(Math.Abs(emitHeight - designThickness * printScale) < 0.002d, "Word-compatible separator rule thickness must scale with the print map.");
    }



    public static void DocxWordCompatibleContinuationRuleSpansDesignBodyWidth()
    {
        // RV06 wclong probe (Word 16.0): word-compatible continuation rules span the
        // full design body mapped once (Office 354.98 emitted), while the renderer
        // emits the shrunk layout body mapped again (269.49 = shrunk times scale).
        // The continuation rule width must be the design body (layout body plus the
        // markup reserve) with the end mark at the rule end.
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:commentRangeStart w:id="0"/></w:r><w:r><w:t xml:space="preserve">commented</w:t></w:r><w:r><w:commentRangeEnd w:id="0"/></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:initials="R" w:date="2024-01-02T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.AllMarkup);
        }
        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, measurer, CancellationToken.None);
        var continuations = layout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(continuations.Length != 0, "Overflowing WC footnotes must place continuation separators.");
        foreach ((DocxLayoutPage page, DocxPlacedRelatedStoryLayout continuation) in continuations)
        {
            double designBody = page.Width - page.MarginLeft - page.MarginRight + page.MarkupMarginReservePoints;
            TestAssert.True(Math.Abs(continuation.SeparatorWidth - designBody) < 0.01d, "WC continuation rule width must be the design body width.");
            TestAssert.True(Math.Abs(continuation.Width - designBody) < 0.01d, "WC continuation story width must cover the design body for the rule clip.");
            if (continuation.TextLines.Count == 1 && continuation.TextLines[0].Segments.Count == 1 &&
                string.IsNullOrWhiteSpace(continuation.TextLines[0].Segments[0].Text))
            {
                double markX = continuation.TextLines[0].Segments[0].X;
                TestAssert.True(Math.Abs(markX - (continuation.X + designBody)) < 0.01d, "WC continuation end mark must sit at the rule end.");
            }
        }
        // Reserve-margin pages keep identity emission over the shrunk body, so the
        // reserve must not join the continuation rule back there.
        DocxLayout reserveLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, measurer, CancellationToken.None);
        var reserveContinuations = reserveLayout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(reserveContinuations.Length != 0, "Overflowing reserve-margin footnotes must place continuation separators.");
        foreach ((DocxLayoutPage reservePage, DocxPlacedRelatedStoryLayout reserveContinuation) in reserveContinuations)
        {
            double shrunkBody = reservePage.Width - reservePage.MarginLeft - reservePage.MarginRight;
            TestAssert.True(Math.Abs(reserveContinuation.SeparatorWidth - shrunkBody) < 0.01d, "Reserve-margin continuation rule width must stay in the shrunk body.");
        }
    }

    public static void DocxLongFootnoteContinuationReservesHeaderZone()
    {
        // RV06 wclong header probes (Word 16.0, h1/h5/nohdr): continued notes leave
        // the static header zone clear (h5 takes 22 vs 25 blind with the rule below
        // all header baselines; nohdr agrees 25/25), while the renderer starts
        // continuation pages at the full page top so the rule lands inside the
        // header text. Every continuation rule must sit below the header baselines
        // of its own page.
        var headerParas = new StringBuilder();
        for (int line = 0; line < 5; line++)
        {
            headerParas.Append("<w:p><w:r><w:t>Header line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">""" + headerParas.ToString() + """</w:hdr>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        var continuations = layout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(continuations.Length != 0, "Overflowing footnotes with headers must place continuation separators.");
        foreach ((DocxLayoutPage page, DocxPlacedRelatedStoryLayout continuation) in continuations)
        {
            double[] headerBaselines = page.StaticTextLines.Select(line => line.BaselineY).ToArray();
            TestAssert.True(headerBaselines.Length != 0, "Continuation pages must carry the static header.");
            double lowestHeaderBaseline = headerBaselines.Min();
            TestAssert.True(continuation.SeparatorY <= lowestHeaderBaseline - 1d, "Continuation rules must sit below the header baselines of their page.");
        }
    }
    public static void DocxLongFootnoteHeadTakeMatchesOfficeWithAptosMetrics()
    {
        // RV06 p1-clamp probes (Word 16.0, h1/b-sweep): the head page takes 24 note
        // lines (2 mixed-size plus long 0..21) with full-width rules below the body,
        // while the renderer takes 23 (long 0..20) with the rule 3.73 low. Synthetic
        // metrics take frame-exact counts that cannot see the real-advance shortfall,
        // so this pins the Office take with installed Aptos advances and skips without
        // a usable Aptos face like other font-environmental tests.
        if (!HasUsableAptosFootnoteFace())
        {
            TestAssert.Skip("Environmental precondition not met: Aptos is not installed.");
        }

        var footnoteParas = new StringBuilder();
        footnoteParas.Append("<w:p><w:r><w:t xml:space=\"preserve\">Footnote twelve opening words</w:t></w:r></w:p>");
        footnoteParas.Append("<w:p><w:r><w:rPr><w:b/><w:sz w:val=\"30\"/><w:szCs w:val=\"30\"/></w:rPr><w:t xml:space=\"preserve\">Footnote fifteen second line</w:t></w:r></w:p>");
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Long note line " + line.ToString(CultureInfo.InvariantCulture) + " with words to fill pages.</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body text with a footnote reference</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">a review note</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> trailing words.</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t xml:space="preserve">Header twelve opening words</w:t></w:r></w:p></w:hdr>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new WindowsFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("Aptos")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        var headLines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .ToArray();
        var longIndexes = headLines
            .Select(line => line.Text)
            .Select(text => text.StartsWith("Long note line ", StringComparison.Ordinal) && int.TryParse(text.Substring(15).Split(' ')[0], out int index) ? index : -1)
            .Where(index => index >= 0)
            .ToArray();
        DocxTextLineLayout firstLine = headLines[0];        string diagnosis = "headLines=" + headLines.Length + ",long=[" + string.Join(",", longIndexes.Take(3)) + ".." + (longIndexes.Length == 0 ? "none" : longIndexes[^1].ToString(CultureInfo.InvariantCulture)) + "],firstLine(h=" + (firstLine.LineHeight?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ",single=" + (firstLine.SingleLineHeight?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ",factor=" + (firstLine.EffectiveLineSpacingFactor?.ToString("F4", CultureInfo.InvariantCulture) ?? "?") + ",after=" + (firstLine.ParagraphAfterSpacing?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ")";
        TestAssert.Equal(24, headLines.Length);
        TestAssert.True(longIndexes.Length == 22 && longIndexes[0] == 0 && longIndexes[^1] == 21, "Head page must take long lines 0..21 like Office. " + diagnosis);
    }

    public static void DocxLongFootnoteHeadBlockSitsAtOfficePosition()
    {
        // RV06 p1-position probes (Word 16.0, h1/b-sweep): with takes exact, the head
        // block must also sit at Office absolute positions (h1-Final rule 682.54 with
        // mixed1 at 664.18), while the renderer seats the whole block ~8 high on
        // capacity slack. Windows-only Aptos pins this like the take test.
        if (!HasUsableAptosFootnoteFace())
        {
            TestAssert.Skip("Environmental precondition not met: Aptos is not installed.");
        }

        var footnoteParas = new StringBuilder();
        footnoteParas.Append("<w:p><w:r><w:t xml:space=\"preserve\">Footnote twelve opening words</w:t></w:r></w:p>");
        footnoteParas.Append("<w:p><w:r><w:rPr><w:b/><w:sz w:val=\"30\"/><w:szCs w:val=\"30\"/></w:rPr><w:t xml:space=\"preserve\">Footnote fifteen second line</w:t></w:r></w:p>");
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Long note line " + line.ToString(CultureInfo.InvariantCulture) + " with words to fill pages.</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body text with a footnote reference</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">a review note</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> trailing words.</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t xml:space="preserve">Header twelve opening words</w:t></w:r></w:p></w:hdr>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new WindowsFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("Aptos")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        DocxLayoutPage page = layout.Pages[0];
        DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxTextLineLayout mixedFirst = page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First(line => line.Text.StartsWith("Footnote twelve opening", StringComparison.Ordinal));
        TestAssert.True(Math.Abs((separator.SeparatorY ?? double.NaN) - 682.54d) < 1.5d, "Head separator rule must sit at the Office position; observed rule bottom=" + (separator.SeparatorY?.ToString(CultureInfo.InvariantCulture) ?? "?") + ".");
        TestAssert.True(Math.Abs(mixedFirst.BaselineY - 664.18d) < 1.5d, "Head footnote content must start at the Office position; observed mixed1 baseline=" + mixedFirst.BaselineY.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static bool HasUsableAptosFootnoteFace()
    {
        try
        {
            FontFaceResolution resolved = new WindowsFontResolver().Resolve(new FontRequest("Aptos"));
            if (resolved.IsFallback)
            {
                return false;
            }

            OpenTypeFont? font = FontProgramLoader.Load(resolved, CancellationToken.None);
            return font is not null && font.HasTrueTypeOutlines;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
        {
            return false;
        }
    }
    public static void DocxFootnoteStoriesTopHangsBelowBodyStart()
    {
        // RV06 p1-clamp probes (Word 16.0, b/ba/2para/bodyb sweeps): Office storiesTop
        // is independent of body size/after/length/before and follows body start with
        // footnote-side moves, while the renderer hung content one body em below the
        // lowest body baseline. Takes come from the same clamped capacity, but the
        // overflowing head block seats bottom-up (storiesTop = margin + takeHeight)
        // instead of leaving capacity slack above the margin.
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutPage page = layout.Pages[0];
        DocxPlacedRelatedStoryLayout body = page.PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        TestAssert.Equal(28, body.TextLines.Count);
        TestAssert.True(Math.Abs(body.TopY - body.Height - page.MarginBottom) < 0.000001d, "The overflowing head block must sit bottom-up on the margin; observed bottom=" + (body.TopY - body.Height).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteSeparatorGapFollowsSingleLineMetrics()
    {
        // RV06 separator-bottom probes (Word 16.0, Times/Aptos/Calibri): the Office
        // footnote gap above the body equals one single-spaced line box minus the
        // first-baseline inset, so it derives from the mark font instead of a fixed
        // 3pt constant. The patched synthetic face has a 1.15em single line, giving
        // a 0.21em gap.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxPlacedRelatedStoryLayout body = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is null);
        double markSize = separator.TextLines.Single().Segments[0].StyleRun.EffectiveProperties.FontSize;
        double separatorBottom = separator.TopY - separator.Height;
        TestAssert.True(Math.Abs(separatorBottom - body.TopY - 0.21d * markSize) < 0.000001d, "Footnote separator gap must follow the single-line box minus the first-baseline inset.");
    }

    public static void DocxFootnoteSeparatorGapKeepsLegacyConstantWithoutLineMetrics()
    {
        // Measurers without single-line metrics keep the legacy 3pt footnote gap.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxPlacedRelatedStoryLayout body = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is null);
        double separatorBottom = separator.TopY - separator.Height;
        TestAssert.True(Math.Abs(separatorBottom - body.TopY - 3d) < 0.000001d, "Legacy footnote separator gap must stay 3pt.");
    }

    public static void DocxLongFootnoteSplitsAcrossPagesWithContinuationSeparator()
    {
        // RV06 footlong probe (Word 16.0): a 60-paragraph footnote splits 24/25/11
        // across 3 pages with full-width continuation rules plus end marks, while the
        // renderer stacks every line on the reference page off-page with no
        // continuation story placed. Continuation separators parse but never place.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        foreach (DocxLayoutPage page in layout.Pages)
        {
            foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
            {
                if (story.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote || story.StoryLayout.Story.Type != null)
                {
                    continue;
                }

                foreach (DocxTextLineLayout line in story.TextLines)
                {
                    TestAssert.True(line.BaselineY >= page.MarginBottom - 1d && line.BaselineY <= page.Height - page.MarginTop + 1d, "Footnote lines must stay inside their page.");
                }
            }
        }
        IReadOnlyList<DocxPlacedRelatedStoryLayout> continuations = layout.Pages.Skip(1).SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator && story.SeparatorY is not null).ToArray();
        TestAssert.True(continuations.Count != 0, "Overflowing footnotes must place continuation separators on continuation pages.");
    }

    public static void DocxLongFootnoteWithoutSeparatorSplitsAcrossPages()
    {
        // RV06 footlong-absent probe (Word 16.0): separator-less footnotes (Word draws
        // default rules) still stack whole stories off-page on overflow. This pins
        // in-bounds lines plus continuation separators for the separator-less path.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        var bodyLines = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).SelectMany(story => story.TextLines).ToArray();
        TestAssert.True(bodyLines.Length == 60, "All separator-less footnote lines must be placed.");
        TestAssert.True(bodyLines.All(line => layout.Pages.Any(page => line.BaselineY >= page.MarginBottom - 1d && line.BaselineY <= page.Height - page.MarginTop + 1d)), "Every separator-less footnote line must sit inside some page.");
        int contRules = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Count(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator && story.SeparatorY is not null);
        TestAssert.True(contRules != 0, "Overflowing separator-less footnotes must place continuation separators.");
    }

    public static void DocxSplitEndnotesKeepEachLineOnExactlyOnePage()
    {
        // RV06 endlong probe (Word 16.0): a 60-paragraph section-end endnote splits
        // 23/25/12 with full-width continuation rules, one text op per line, while the
        // renderer triplicates every line across 3 slice pages (180 ops for 60 lines)
        // because slices shift without filtering and only clip at emission.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.True(layout.Pages.Count > 1, "Overflowing endnotes must paginate.");
        var lineKeys = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type != DocxRelatedStoryType.Separator && story.StoryLayout.Story.Type != DocxRelatedStoryType.ContinuationSeparator).SelectMany(story => story.TextLines).Select(line => (line.SourceBlockIndex, line.SourceParagraphIndex, line.SourceLineIndex)).ToArray();
        TestAssert.Equal(lineKeys.Distinct().Count(), lineKeys.Length);
    }

    public static void DocxEndnoteSeparatorIsPlacedWithRuleMark()
    {
        // RV06 endnote probes: Office draws the endnote separator rule with a mark
        // space at the rule end, like footnotes.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        TestAssert.Equal(1, separatorLine.Segments.Count);
        TestAssert.Equal(" ", separatorLine.Segments[0].Text);
        TestAssert.Equal(144d, separatorLine.Segments[0].X - separatorLine.X);
    }

    public static void DocxDocumentEndnoteSeparatorIsPlacedWithRuleMark()
    {
        // RV06 endnote probes: document-end endnotes draw the separator rule with a
        // mark space at the rule end, like section-end endnotes.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        TestAssert.Equal(1, separatorLine.Segments.Count);
        TestAssert.Equal(" ", separatorLine.Segments[0].Text);
        TestAssert.Equal(144d, separatorLine.Segments[0].X - separatorLine.X);
    }

    public static void DocxFootnoteRemainderAccountsSeparatorUntilPlaced()
    {
        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "31",
            [],
            [],
            [], null);
        var storyLayout = new DocxRelatedStoryLayout(story, 0, [], [], [], [], 0d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(
            0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "31", null, "1", 0, 1, 4));
        var inFlight = new DocxLayoutEngine.InFlightRelatedStory(storyLayout, null, null, 0, location);
        double[] lineBoxes = [10d, 10d, 10d, 10d, 10d];

        double remainder = DocxLayoutEngine.FootnoteRemainderHeight(inFlight, lineBoxes, null);

        TestAssert.Equal(53d, remainder);
    }

    public static void DocxFootnoteRemainderShrinksWithPlacedTakes()
    {
        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "32",
            [],
            [],
            [], null);
        DocxTextRun lineRun = new("Shared note", 10d, null, false, false, false, null, null);
        DocxTextLineLayout RemainderLine(int index)
        {
            return new DocxTextLineLayout("Shared note line " + index.ToString(CultureInfo.InvariantCulture), lineRun, 10d, 72d, 700d - (12d * index), 200d, [], null, null, null, null, 12d, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        }

        var storyLayout = new DocxRelatedStoryLayout(story, 0, [RemainderLine(0), RemainderLine(1), RemainderLine(2), RemainderLine(3), RemainderLine(4)], [], [], [], 60d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(
            0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "32", null, "1", 0, 1, 4));
        var inFlight = new DocxLayoutEngine.InFlightRelatedStory(storyLayout, null, null, 0, location)
        {
            PlacedLineCount = 2,
            SeparatorPlaced = true
        };
        double[] lineBoxes = [10d, 10d, 10d, 10d, 10d];

        double remainder = DocxLayoutEngine.FootnoteRemainderHeight(inFlight, lineBoxes, null);

        TestAssert.Equal(30d, remainder);
        TestAssert.Equal(3, inFlight.RemainingLineCount);
        TestAssert.Equal(30d, inFlight.RemainingContentHeight(lineBoxes));
    }

    public static void DocxNarrowedFootnoteRemainderRebasesAtStoryOrigin()
    {
        DocxTextRun run = new("Shared note", 10d, null, false, false, false, null, null);
        DocxTextLineLayout Line(string text, double baseline)
        {
            return new DocxTextLineLayout(text, run, 10d, 72d, baseline, 200d, [], null, null, null, null, 12d, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        }

        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "33",
            [],
            [],
            [], null);
        var storyLayout = new DocxRelatedStoryLayout(
            story, 0, [Line("Shared note line zero", 700d), Line("Shared note line one", 688d), Line("Shared note line two", 676d)], [], [], [], 36d);

        DocxRelatedStoryLayout narrowed = DocxLayoutEngine.NarrowStoryTextLinesForOffset(storyLayout, 1);

        TestAssert.Equal(2, narrowed.TextLines.Count);
        TestAssert.Equal(700d, narrowed.TextLines[0].BaselineY);
        TestAssert.Equal(688d, narrowed.TextLines[1].BaselineY);
        TestAssert.Equal(24d, narrowed.ContentHeight);
    }

    public static void DocxSharedFootnotePagesKeepBodyAboveNotes()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 40, fillerParagraphCount: 60, footnoteId: "41");
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(40, placedTexts.Length);
        TestAssert.Equal(40, placedTexts.Distinct().Count());

        foreach (DocxLayoutPageSnapshot page in snapshot.Pages)
        {
            if (page.Items.Count == 0 || page.PlacedFootnoteStoryCount == 0)
            {
                continue;
            }

            double bodyBottom = page.Items.Min(item => item.Y);
            double footnoteTop = page.PlacedRelatedItems.Max(item => item.Y + item.Height);
            TestAssert.True(bodyBottom >= footnoteTop, "Shared pages must keep body above placed notes.");
        }
    }

    public static void DocxExhaustedFootnoteDrainUsesDedicatedPages()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 200, fillerParagraphCount: 0, footnoteId: "42");
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.True(layout.Pages.Count > 2, "The long note must span several pages.");
        TestAssert.Equal(200, placedTexts.Length);
        TestAssert.Equal(200, placedTexts.Distinct().Count());
        TestAssert.Equal(0, layout.Pages[^1].Items.Count);
        TestAssert.True(layout.Pages[^1].PlacedRelatedStories.Count != 0, "The exhausted drain must append dedicated note pages.");
    }

    public static void DocxMidBodyFootnoteAnchorKeepsPrecedingFillersOnPage()
    {
        const int fillersBefore = 30;
        DocxDocument document = CreateMidBodyInterleavingDocument(fillersBefore, noteParagraphCount: 60, footnoteId: "43", fillersAfter: 30);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        int PageWithBlock(int blockIndex)
        {
            for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
            {
                foreach (DocxLayoutItem item in layout.Pages[pageIndex].Items)
                {
                    if (item is DocxTextLineLayout line && line.SourceBlockIndex == blockIndex)
                    {
                        return pageIndex;
                    }
                }
            }

            return -1;
        }

        TestAssert.True(PageWithBlock(fillersBefore - 1) >= 0, "Preceding fillers must be placed.");
        TestAssert.Equal(PageWithBlock(fillersBefore - 1), PageWithBlock(fillersBefore));

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(60, placedTexts.Length);
        TestAssert.Equal(60, placedTexts.Distinct().Count());
    }

    public static void DocxTableAnchoredFootnoteSharesBodyPages()
    {
        DocxParagraph cellAnchor = DocxTests.CreateDocxLayoutParagraph("Cell note marker", 10d, 12d) with
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
                    TextOffsetInRun: 4)
            ]
        };
        var cell = new DocxTableCell(
            string.Empty,
            [cellAnchor, DocxTests.CreateDocxLayoutParagraph("Cell trailing line", 10d, 12d)],
            null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], 30d)]);
        DocxParagraphElement[] noteElements = Enumerable.Range(0, 120)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "51",
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement> { };
        var paragraphs = new List<DocxParagraph> { cellAnchor };
        for (int index = 0; index < 60; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

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
            new DocxBodyElement[] { new DocxTableElement(table) }.Concat(bodyElements).ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = [footnoteStory]
        };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        bool sharedPageExists = false;
        foreach (DocxLayoutPage page in layout.Pages)
        {
            bool hasFiller = page.Items.OfType<DocxTextLineLayout>().Any(line => line.Text.StartsWith("Filler body line", StringComparison.Ordinal));
            bool hasTable = page.Items.OfType<DocxTableRowLayout>().Any() || page.Items.OfType<DocxTextLineLayout>().Any(line => line.Text.StartsWith("Cell ", StringComparison.Ordinal));
            bool hasNote = page.PlacedRelatedStories
                .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
                .SelectMany(story => story.TextLines)
                .Any();
            if (hasFiller && hasNote && !hasTable)
            {
                sharedPageExists = true;
            }
        }

        TestAssert.True(sharedPageExists, "Table-anchored notes must share body pages instead of dedicating continuation pages.");

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(120, placedTexts.Length);
        TestAssert.Equal(120, placedTexts.Distinct().Count());
    }

    public static void DocxKeepChainBlockIndexesFollowKeepNext()
    {
        DocxParagraph Keep(string text)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d, keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        }

        DocxParagraph plain = DocxTests.CreateDocxLayoutParagraph("Plain", 10d, 12d);
        var table = new DocxTable(null, [60d], [new DocxTableRow([new DocxTableCell("Cell", [plain], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d)]);
        IReadOnlyList<DocxBodyElement> elements = [new DocxParagraphElement(Keep("Keep zero")), new DocxParagraphElement(Keep("Keep one")), new DocxParagraphElement(plain), new DocxTableElement(table), new DocxParagraphElement(Keep("Keep four"))];

        TestAssert.True(new List<int> { 0, 1, 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 0)), "Keep chain from 0 must span both keep paras plus the plain follower.");
        TestAssert.True(new List<int> { 1, 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 1)), "Keep chain from 1 must span its keep para plus the plain follower.");
        TestAssert.True(new List<int> { 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 2)), "Plain paragraphs carry no chain.");
        TestAssert.True(new List<int> { 3 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 3)), "Table elements carry no chain.");
        TestAssert.True(new List<int> { 4 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 4)), "Trailing keep with no follower carries no chain.");
    }

    public static void DocxSharedContinuedFootnotePagesCarryContinuationRules()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 120, fillerParagraphCount: 60, footnoteId: "44", includeSeparators: true);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        bool headPageSeen = false;
        foreach (DocxLayoutPage page in layout.Pages)
        {
            bool hasNormalLines = page.PlacedRelatedStories
                .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
                .SelectMany(story => story.TextLines)
                .Any();
            if (!hasNormalLines)
            {
                continue;
            }

            if (!headPageSeen)
            {
                headPageSeen = true;
                continue;
            }
        TestAssert.True(
                page.PlacedRelatedStories.Any(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator),
                "Continued footnote pages must open with a continuation rule.");

        }

    }

    private static DocxDocument CreateMidBodyInterleavingDocument(int fillersBefore, int noteParagraphCount, string footnoteId, int fillersAfter)
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    footnoteId,
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraphElement[] noteElements = Enumerable.Range(0, noteParagraphCount)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            footnoteId,
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement>();
        var paragraphs = new List<DocxParagraph>();
        for (int index = 0; index < fillersBefore; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        bodyElements.Add(new DocxParagraphElement(anchor));
        paragraphs.Add(anchor);
        for (int index = 0; index < fillersAfter; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + (fillersBefore + index).ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        return new DocxDocument(
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
            bodyElements.ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = [footnoteStory]
        };
    }

    private static DocxDocument CreateInterleavingDocument(int noteParagraphCount, int fillerParagraphCount, string footnoteId, bool includeSeparators = false)
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    footnoteId,
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraphElement[] noteElements = Enumerable.Range(0, noteParagraphCount)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            footnoteId,
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement> { new DocxParagraphElement(anchor) };
        DocxRelatedStory? separatorStory = null;
        DocxRelatedStory? continuationStory = null;
        if (includeSeparators)
        {
            separatorStory = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "0", [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Note separator mark", 10d, 12d))], [], [], DocxRelatedStoryType.Separator);
            continuationStory = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "1", [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Note continuation mark", 10d, 12d))], [], [], DocxRelatedStoryType.ContinuationSeparator);
        }

        var paragraphs = new List<DocxParagraph> { anchor };
        for (int index = 0; index < fillerParagraphCount; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        return new DocxDocument(
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
            bodyElements.ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = separatorStory is null || continuationStory is null ? [footnoteStory] : [footnoteStory, separatorStory, continuationStory]
        };
    }

    public static void DocxEndnoteSeparatorIgnoresDirectParagraphSpacing()
    {
        // RV06 endnote-spacing probes (Word 16.0, edge-endsepsp/endsepspb): Office
        // holds document-end endnotes byte-identical across separator before/after
        // 0 vs 24pt, so separator stories lay out with latent-default spacing while
        // the renderer flows direct values into the separator height (24pt error).
        double topAfter0 = LayoutEndnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double topAfter24 = LayoutEndnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double topDefault = LayoutEndnoteBodyTopWithSeparatorSpacing(string.Empty);
        TestAssert.True(Math.Abs(topAfter0 - topAfter24) < 0.000001d, "Endnote separator after-spacing must not move the note block; observed shift=" + Math.Abs(topAfter0 - topAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(topAfter0 - topDefault) < 0.000001d, "Endnote separator after-spacing must match latent defaults; observed shift=" + Math.Abs(topAfter0 - topDefault).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteSeparatorIgnoresDirectParagraphSpacing()
    {
        // Shared story-layout path with endnote separators (Office evidence is
        // endnote-only; no diverging footnote evidence): direct separator spacing
        // must not move footnote content, and the separator rule must not ride the
        // direct spacing either (the footnote content hangs off the stories top
        // while the rule sits in the separator story, so the rule is the moving
        // part here).
        double topAfter0 = LayoutFootnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double topAfter24 = LayoutFootnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double topDefault = LayoutFootnoteBodyTopWithSeparatorSpacing(string.Empty);
        TestAssert.True(Math.Abs(topAfter0 - topAfter24) < 0.000001d, "Footnote separator after-spacing must not move the note block; observed shift=" + Math.Abs(topAfter0 - topAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(topAfter0 - topDefault) < 0.000001d, "Footnote separator after-spacing must match latent defaults; observed shift=" + Math.Abs(topAfter0 - topDefault).ToString(CultureInfo.InvariantCulture) + ".");
        double ruleAfter0 = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double ruleAfter24 = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double ruleDefault = LayoutFootnoteSeparatorRuleWithSpacing(string.Empty);
        TestAssert.True(Math.Abs(ruleAfter0 - ruleAfter24) < 0.000001d, "Footnote separator rule must not ride direct after-spacing; observed shift=" + Math.Abs(ruleAfter0 - ruleAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(ruleAfter0 - ruleDefault) < 0.000001d, "Footnote separator rule must match latent defaults; observed shift=" + Math.Abs(ruleAfter0 - ruleDefault).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxEndnoteSeparatorIgnoresExactLineSpacing()
    {
        // RV06 exact-line probe (Word 16.0, edge-endsepspex): Office shifts the
        // endnote block identically with exact-24 and auto separator lines
        // (minus 8.76 both), so separator stories also ignore exact line rules
        // while the renderer clamps the mark line box to 24pt (shift minus 3.76).
        double topExact = LayoutEndnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:line=\"480\" w:lineRule=\"exact\"/></w:pPr>");
        double topAuto = LayoutEndnoteBodyTopWithSeparatorSpacing(string.Empty);
        TestAssert.True(Math.Abs(topExact - topAuto) < 0.000001d, "Endnote separator exact line spacing must match auto lines; observed shift=" + Math.Abs(topExact - topAuto).ToString(CultureInfo.InvariantCulture) + ".");
    }
    public static void DocxEndnoteSeparatorGapUsesMarkMetrics()
    {
        // RV06 height-model probes (Word COM references edge-endsepheight-* plus the
        // 10/12/14pt mark grids): document-end content placement below the separator
        // grows with mark size (Office rule-to-first slope carries a size-driven
        // below-gap), while the renderer drops a legacy constant 3pt. A synthetic
        // singleLineEm of 1.5em makes the mark-metrics gap (1.5 minus 0.94) times the
        // 10pt mark size, pinning 5.6pt against the legacy constant. Pre-fix the
        // separator-to-content distance renders at 3pt.
        double gap = LayoutEndnoteSeparatorContentGapWithMarkSize(20, 1.5d);
        TestAssert.True(Math.Abs(gap - 5.6d) < 0.000001d, "Document-end below-separator gap must follow mark metrics.");
    }
    public static void DocxEndnoteSeparatorHeightGrowsWithFirstInset()
    {
        // RV06 height-model probes (empty/exact-24/text separator variants place
        // identically in Word): Office separator placement ignores laid-out line
        // boxes, growing only by the first-line inset with mark size. The laid-out
        // box slope would separate 10pt and 20pt marks by a full line box per 10pt;
        // the size-driven placement height separates them by one first inset.
        // Pre-fix the placed heights differ by the laid-out box slope.
        double shortHeight = LayoutEndnoteSeparatorHeightWithMarkSize(20);
        double tallHeight = LayoutEndnoteSeparatorHeightWithMarkSize(40);
        TestAssert.True(Math.Abs((tallHeight - shortHeight) - 9.4d) < 0.000001d, "Document-end separator placement height must grow by one first inset per 10pt.");
    }

    public static void DocxEndnoteSeparatorHeightCorrectsSingleLineDeficit()
    {
        // RV06 four-family separator grids (Word COM references edge-endsepgrid-cal/tmr/tah/vdn):
        // Office separator height carries a singleLine-deficit term against the validated
        // Calibri anchor (measured singleLineEm 1.2207): Times, Tahoma and Verdana share 1.15
        // yet place about 0.7pt lower at 12pt, which no laid-out box spread explains. A synthetic
        // 1.15em mark must stand 0.84 times the 0.0707 deficit times 12pt above the anchored
        // height. Pre-fix both heights are equal.
        double calHeight = LayoutEndnoteSeparatorHeightWithSingleLineEm(24, 1.2207d);
        double lowHeight = LayoutEndnoteSeparatorHeightWithSingleLineEm(24, 1.15d);
        TestAssert.True(Math.Abs((lowHeight - calHeight) - 0.71d) < 0.01d, "Document-end separator height must correct the singleLine deficit; observed shift=" + (lowHeight - calHeight).ToString(CultureInfo.InvariantCulture) + ".");
    }
    private static double LayoutEndnoteSeparatorHeightWithSingleLineEm(int docDefaultsHalfPoints, double singleLineEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SingleLineEmTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), singleLineEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }
    public static void DocxEndnoteAbsentSeparatorNormalizesToDefault()
    {
        // RV06 height-model probes (Word COM reference edge-endsepheight-absent):
        // Word normalizes separator-less endnote parts with a default separator
        // (absent renders identically to mark: same rule, same placement), while
        // the renderer skips placement entirely (content a full separator higher
        // with no rule). Pre-fix no separator story is placed. Footnote parts
        // keep legacy absent behavior.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        TestAssert.True(separator.Height > 0d, "Synthesized default separator must take vertical space.");
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        TestAssert.True(Math.Abs(((separator.TopY - separator.Height) - content.TopY) - 3d) < 0.000001d, "Default separator content must keep the legacy gap with the FamilyWidth measurer.");
    }

    private sealed class SingleLineEmTextMeasurer(IDocxTextMeasurer inner, double singleLineEm) : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = singleLineEm;
            return true;
        }
    }

    private static double LayoutEndnoteSeparatorContentGapWithMarkSize(int docDefaultsHalfPoints, double singleLineEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SingleLineEmTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), singleLineEm), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }
    private static double LayoutEndnoteSeparatorHeightWithMarkSize(int docDefaultsHalfPoints)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }

    public static void DocxEndnoteSeparatorHeightAddsSmallAscSupplement()
    {
        double control = LayoutEndnoteSeparatorHeightWithWindowsAscender(16, 0.9521d);
        double small = LayoutEndnoteSeparatorHeightWithWindowsAscender(16, 0.89d);
        TestAssert.True(Math.Abs((small - control) - 0.50d) < 0.01d, "Document-end separator height must add the small ascender supplement; observed shift=" + (small - control).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class WindowsAscenderTextMeasurer(IDocxTextMeasurer inner, double windowsAscenderEm) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => windowsAscenderEm * fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;
    }

    private static double LayoutEndnoteSeparatorHeightWithWindowsAscender(int docDefaultsHalfPoints, double windowsAscenderEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new WindowsAscenderTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), windowsAscenderEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }

    public static void DocxFootnoteFirstBaselineIgnoresLaterRunFonts()
    {
        // RV05 fnmix probes (Word COM references edge-fnmix/edge-fnmix3: note first
        // baselines sit at 85.46 regardless of run order or family): related-story
        // content keeps legacy widest-run insets, so a trailing Tahoma run must not
        // move the first baseline. Pre-scoping (max-hhea) it drops by 0.73.
        double control = LayoutFootnoteFirstBaselineWithSecondRun("Calibri");
        double mixed = LayoutFootnoteFirstBaselineWithSecondRun("Tahoma");
        TestAssert.True(Math.Abs(mixed - control) < 0.000001d, "Footnote first baseline must ignore later-run fonts; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + " control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class StoryHheaTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize)
        {
            double em = run?.FontFamily == "Tahoma" ? 1.0005d : 0.75d;
            return em * fontSize;
        }
    }

    private static double LayoutFootnoteFirstBaselineWithSecondRun(string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note start words </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="FAM" w:hAnsi="FAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">middle mixed words</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("FAM", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new StoryHheaTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First(line => line.Text.StartsWith("Note start words", StringComparison.Ordinal)).BaselineY;
    }
    public static void DocxFootnoteLineHeightUsesMaxAcrossRuns()
    {
        double control = LayoutFootnotePitchWithRuns("Calibri", "Calibri");
        double tall = LayoutFootnotePitchWithRuns("Tahoma", "Tahoma");
        double mixed = LayoutFootnotePitchWithRuns("Calibri", "Tahoma");
        double swapped = LayoutFootnotePitchWithRuns("Tahoma", "Calibri");
        TestAssert.True(Math.Abs(mixed - tall) < 0.000001d, "Footnote line height must use max across runs; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + " tall=" + tall.ToString(CultureInfo.InvariantCulture) + " control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(swapped - tall) < 0.000001d, "Footnote line height must ignore run order; swapped=" + swapped.ToString(CultureInfo.InvariantCulture) + " tall=" + tall.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class StoryHeightTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 20d : 10d;
    }

    private static double LayoutFootnotePitchWithRuns(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new StoryHeightTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxEndnoteSeparatorTextMatchesMarkEmission()
    {
        // RV06 height-model probes (Word COM reference edge-endsepheight-text):
        // Office drops non-mark separator text while drawing the rule (text variant
        // renders the same 12 ops as the mark variant with no dashes); the renderer
        // emits the dashes as extra text shows. Pre-fix the text variant shows more.
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        var resolver = new DocxTests.SingleResolutionFontResolver(font.Value.Resolution);
        int markShows = CountSeparatorVariantTextShows("<w:p><w:r><w:separator/></w:r></w:p>", resolver);
        int textShows = CountSeparatorVariantTextShows("<w:p><w:r><w:t xml:space=\"preserve\">---</w:t></w:r></w:p>", resolver);
        TestAssert.Equal(markShows, textShows);
    }

    private static int CountSeparatorVariantTextShows(string separatorParagraph, DocxTests.SingleResolutionFontResolver resolver)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0">SEPPARA</w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>""".Replace("SEPPARA", separatorParagraph)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        PdfPage page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();
        return DocxTests.CountPdfTextShows(page.Content);
    }
    public static void DocxFootnoteSeparatorIgnoresExactLineSpacing()
    {
        // Shared story-layout path with endnote separators (Office evidence is
        // endnote-only; no diverging footnote evidence): the footnote rule rides
        // the separator bottom, so an exact line rule that Office ignores would
        // move our rule with the taller story.
        double ruleExact = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:line=\"480\" w:lineRule=\"exact\"/></w:pPr>");
        double ruleAuto = LayoutFootnoteSeparatorRuleWithSpacing(string.Empty);
        TestAssert.True(Math.Abs(ruleExact - ruleAuto) < 0.000001d, "Footnote separator exact line spacing must match auto lines; observed shift=" + Math.Abs(ruleExact - ruleAuto).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutEndnoteBodyTopWithSeparatorSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).TopY;
    }

    private static double LayoutFootnoteBodyTopWithSeparatorSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).TopY;
    }

    private static double LayoutFootnoteSeparatorRuleWithSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null).SeparatorY ?? 0d;
    }
    }
