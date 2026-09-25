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
            .Single(story => story.Kind == "Footnote" && story.Id == "16");
        PdfPage placedPage = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single(page => page.Content.Contains(" re W n", StringComparison.Ordinal));

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
    }}
