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

internal static class DocxCommentsTests
{
    public static void DocxAllMarkupRendererDrawsCommentBalloons()
    {
        string input = DocxTests.WriteCommentMarkerProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Contains("0.5 w", page.Content);
        TestAssert.Contains(" re B*", page.Content);
        TestAssert.Contains(" l S", page.Content);
    }

    public static void DocxAllMarkupRendererPlacesStaticStoryCommentBalloons()
    {
        DocxParagraph headerParagraph = DocxTests.CreateCommentMarkerParagraph("Header review anchor", "1");
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [headerParagraph]
            }
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Header comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            240d,
            200d,
            10d,
            80d,
            10d,
            10d,
            pageSettings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);

        TestAssert.True(
            placements.Any(placement => placement.Kind == "Comment"),
            "All-markup comment balloons should be anchored from rendered static header/footer story lines.");
    }

    public static void DocxAllMarkupRendererPlacesRelatedStoryCommentBalloons()
    {
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
        DocxParagraph footnoteParagraph = DocxTests.CreateCommentMarkerParagraph("Footnote review anchor", "2");
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "9",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "2",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Footnote comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [footnoteStory, commentStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);
        DocxMarkupBalloonPlacementSnapshot commentPlacement = placements.Single(placement => placement.Kind == "Comment");

        TestAssert.True(commentPlacement.AnchorConnectorX >= document.MarginLeftPoints, "All-markup comment balloons should be anchored from placed footnote/endnote story lines in page coordinates.");

        DocxParagraph endnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "11",
                    CustomMarkFollowsValue: null,
                    DisplayText: "i",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateCommentMarkerParagraph("Endnote review anchor", "12");
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "11",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxRelatedStory endnoteCommentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "12",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Endnote comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument endnoteDocument = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(endnoteAnchor)], [])
            with
            {
                RelatedStories = [endnoteStory, endnoteCommentStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };

        DocxMarkupBalloonPlacementSnapshot endnoteCommentPlacement = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(endnoteDocument)
            .Single(placement => placement.Kind == "Comment");

        TestAssert.True(endnoteCommentPlacement.AnchorConnectorX >= endnoteDocument.MarginLeftPoints, "All-markup comment balloons should be anchored from placed endnote story lines in page coordinates.");
    }

    public static void DocxAllMarkupRendererPlacesRelatedStoryRevisionBalloons()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "10",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote revised anchor", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "10", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, [])
            ]
        };
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "10",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [footnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);
        DocxMarkupBalloonPlacementSnapshot revisionPlacement = placements.Single(placement => placement.Kind == "Revision");

        TestAssert.Equal(1, revisionPlacement.RevisionCandidateCount);
        TestAssert.True(revisionPlacement.AnchorConnectorX >= document.MarginLeftPoints, "All-markup revision balloons should be anchored from placed footnote/endnote story lines in page coordinates.");

        DocxParagraph endnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "11",
                    CustomMarkFollowsValue: null,
                    DisplayText: "i",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote revised anchor", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "11", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, [])
            ]
        };
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "11",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxDocument endnoteDocument = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(endnoteAnchor)], [])
            with
            {
                RelatedStories = [endnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };

        DocxMarkupBalloonPlacementSnapshot endnoteRevisionPlacement = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(endnoteDocument)
            .Single(placement => placement.Kind == "Revision");

        TestAssert.Equal(1, endnoteRevisionPlacement.RevisionCandidateCount);
        TestAssert.True(endnoteRevisionPlacement.AnchorConnectorX >= endnoteDocument.MarginLeftPoints, "All-markup revision balloons should be anchored from placed endnote story lines in page coordinates.");
    }

    public static void DocxAllMarkupRendererPlacesFloatingTextBoxMarkupBalloons()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateCommentMarkerParagraph("Text box review anchor", "5") with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "5", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, [])
            ]
        };
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "5",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Text box comment body", 10d, 12d))],
            [],
            [], null);
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
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .ToArray();

        TestAssert.True(
            placements.Any(placement => placement.Kind == "Markup" && placement.CandidateCount >= 2),
            "All-markup comment and revision balloon candidates should be anchored from floating text-box lines before same-anchor grouping.");
    }

    public static void DocxAllMarkupRendererPlacesStaticFloatingTextBoxMarkupBalloons()
    {
        DocxParagraph textBoxParagraph = DocxTests.CreateCommentMarkerParagraph("Static text box review anchor", "6") with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "6", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, [])
            ]
        };
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)])]
            }
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "6",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Static text box comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            300d,
            300d,
            30d,
            90d,
            30d,
            30d,
            pageSettings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .ToArray();

        TestAssert.True(
            placements.Any(placement => placement.Kind == "Markup" && placement.CandidateCount >= 2),
            "All-markup comment and revision balloon candidates should be anchored from static floating text-box lines before same-anchor grouping.");
    }

    public static void DocxWordCompatibleAllMarkupOmitsInlineCommentMarkerLabels()
    {
        string input = DocxTests.WriteCommentMarkerProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        PdfPage preservePage = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();
        PdfPage wordPage = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(
            DocxTests.CountPdfTextShows(wordPage.Content) < DocxTests.CountPdfTextShows(preservePage.Content),
            "Word-compatible all-markup should keep comment marker graphics without emitting inline marker label glyphs.");
    }

    public static void DocxWordCompatibleAllMarkupRendersBareCommentReferenceMarker()
    {
        string input = DocxTests.WriteCommentMarkerProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0.82 0.204 0.22 RG", page.Content);
        TestAssert.Contains("0.973 0.863 0.867 rg", page.Content);
        int strokeCount = DocxTests.CountOccurrences(page.Content, " l S");
        TestAssert.True(
            strokeCount >= 6,
            "Word-compatible all-markup should draw a point marker for bare commentReference anchors without explicit range bounds; observed " + strokeCount + " stroke-line operations.");
        TestAssert.True(
            !page.Content.Contains("1 0.753 0 rg", StringComparison.Ordinal),
            "Word-compatible all-markup should not fall back to legacy yellow marker boxes for bare commentReference anchors.");
    }

    public static void DocxWordCompatibleAllMarkupRendersCommentRangeBrackets()
    {
        const string commentText = "Review note";
        DocxTextRun commentRun = new(commentText, 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        DocxParagraph paragraph = new(
            [commentRun],
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
                new DocxCommentRange("1", 0, 0, 1, commentText.Length, 1, 0)
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

        TestAssert.Contains("0.82 0.204 0.22 RG", page.Content);
        TestAssert.Contains("0.973 0.863 0.867 rg", page.Content);
        TestAssert.True(
            DocxTests.CountOccurrences(page.Content, " l S") >= 6,
            "Word-compatible all-markup should draw comment range brackets as stroke segments.");
        TestAssert.True(
            !page.Content.Contains("1 0.753 0 rg", StringComparison.Ordinal),
            "Word-compatible all-markup should not draw legacy yellow comment marker boxes.");
    }

    public static void DocxWordCompatibleAllMarkupRendersCommentRangeBracketsAcrossTextFlows()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateCommentRangeParagraph("Body range", "1");
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(bodyParagraph)], [], DocxPageSettings.Empty, []) with
            {
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "body text");

        DocxParagraph tableParagraph = DocxTests.CreateCommentRangeParagraph("Table range", "2");
        var tableCell = new DocxTableCell(
            "Table range",
            [tableParagraph],
            FillHex: null,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders: [],
            DocxTableCellMargins.Empty);
        var table = new DocxTable(
            LayoutValue: null,
            ColumnWidthsPoints: [100d],
            Rows: [new DocxTableRow([tableCell], HeightPoints: 18d)]);
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxTableElement(table)], [], DocxPageSettings.Empty, []) with
            {
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "table cell text");

        DocxParagraph headerParagraph = DocxTests.CreateCommentRangeParagraph("Header range", "3");
        DocxPageSettings headerSettings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(headerParagraph)]
            }
        };
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([], [], headerSettings, []) with
            {
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "static header text");

        DocxParagraph drawingAnchor = DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateCommentRangeParagraph("Text box range", "4");
        var floatingDrawing = new DocxFloatingDrawing(
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
            VerticalOffsetValue: "3657600",
            WrapKind: DocxFloatingWrapKind.None,
            WrapTextValue: null,
            SourceParagraphIndex: 0,
            SourceBlockIndex: 0, ImageRelationshipId: null, Image: null)
        {
            TextBoxBodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(drawingAnchor)], [floatingDrawing], DocxPageSettings.Empty, []) with
            {
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "floating text-box text");

        DocxParagraph staticTextBoxParagraph = DocxTests.CreateCommentRangeParagraph("Static text box range", "5");
        DocxPageSettings staticFloatingSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(staticTextBoxParagraph)])]
            }
        };
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([], [], staticFloatingSettings, []) with
            {
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "static floating text-box text");

        DocxParagraph footnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "21",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph footnoteRangeParagraph = DocxTests.CreateCommentRangeParagraph("Footnote range", "6");
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "21",
            [new DocxParagraphElement(footnoteRangeParagraph)],
            [],
            [], null);
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(footnoteAnchor)], [], DocxPageSettings.Empty, []) with
            {
                RelatedStories = [footnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "placed footnote text");

        DocxParagraph endnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "22",
                    CustomMarkFollowsValue: null,
                    DisplayText: "i",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph endnoteRangeParagraph = DocxTests.CreateCommentRangeParagraph("Endnote range", "7");
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "22",
            [new DocxParagraphElement(endnoteRangeParagraph)],
            [],
            [], null);
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(endnoteAnchor)], [], DocxPageSettings.Empty, []) with
            {
                RelatedStories = [endnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "placed endnote text");
    }

    public static void DocxWordCompatibleAllMarkupRendersWrappedCommentRangeContinuationMarkers()
    {
        const string startText = "Start ";
        const string continuedText = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega alpha beta gamma delta epsilon";
        DocxTextRun startRun = new(startText, 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        DocxTextRun continuedRun = new(continuedText, 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 1,
            SourceTextOffsetInRun = 0
        };
        DocxParagraph paragraph = new(
            [startRun, continuedRun],
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
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 2, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ],
            CommentRanges =
            [
                new DocxCommentRange("1", 0, 0, 2, 0, 2, 0)
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
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxTextEmissionSnapshot textEmission = renderer.InspectTextEmission(document);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(
            textEmission.Lines.Count(line => !line.IsStaticStory && line.TextLength != 0) >= 2,
            "The public comment range fixture should wrap across multiple body lines.");
        TestAssert.True(
            DocxTests.CountOccurrences(page.Content, " 11.625 re f") >= 2,
            "Word-compatible all-markup should draw comment range fill markers on wrapped continuation lines, not only on the line containing the range start.");
    }

    public static void DocxWordCompatibleAllMarkupRendersGroupedCommentTextWithOfficeBalloonFont()
    {
        DocxParagraph boldSeedParagraph = new(
            [new DocxTextRun("Bold label seed", 10d, null, true, false, false, null, null)],
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
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Grouped anchor text", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 0, RunChildIndex: 0, TextOffsetInRun: 8, DisplayText: null)
            ],
            Revisions =
            [
                // Office A/B: Word shows insertions inline, so the grouped revision candidate below
                // is a formatting revision (the only revision kind Word balloons).
                new DocxRevisionInfo(DocxRevisionKind.RunPropertiesChange, "1", "Reviewer", "2026-06-01T00:00:00Z", "rPrChange", null, propertyElementNames: ["b"])
            ]
        };
        DocxParagraph commentParagraph = DocxTests.CreateDocxLayoutParagraph("Public comment alpha beta gamma delta epsilon for wrapping", 10d, 12d);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(commentParagraph)],
            [],
            [], null)
        {
            CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", "2026-06-01T00:00:00Z", null, null, null, null)
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
            [new DocxParagraphElement(boldSeedParagraph), new DocxParagraphElement(paragraph)],
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

        DocxMarkupBalloonPlacementSnapshot[] placements = renderer.InspectMarkupBalloons(document).ToArray();
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        DocxMarkupBalloonPlacementSnapshot mixedPlacement = placements.Single(placement => placement.Kind == "Markup");
        TestAssert.True(
            mixedPlacement.CandidateCount == 2 && mixedPlacement.CommentCandidateCount == 1 && mixedPlacement.RevisionCandidateCount == 1,
            "Same-anchor comment and revision candidates should retain grouped private-safe geometry and candidate classes.");
        TestAssert.True(
            mixedPlacement.BodySummaryPartCount == 2 && mixedPlacement.WordCompatibleBodySummaryPartCount == 2,
            "Grouped comment/revision balloons should retain both private-safe summary parts in Word-compatible body text.");
        TestAssert.True(
            DocxTests.CountOccurrences(page.Content, "8.203 Tf") >= 3,
            "Word-compatible grouped comment balloons should render Office-sized title and wrapped body text.");
        // Title TJ adjustments are point-space tracking over the balloon font: -0.03357 * 1000 / 8.20253
        // (9pt design over this document lane-fit scale 612 / 671.5), i.e. -4.09266, which "0.###"
        // rounds to -4.093 (was -4.813 at 6.975pt).
        TestAssert.True(
            page.Content.Contains("-4.093", StringComparison.Ordinal),
            "Word-compatible grouped comment balloon titles should use positioned glyph advances.");
        TestAssert.True(
            DocxTests.CountOccurrences(page.Content, "/F3 8.203 Tf") >= 1 && DocxTests.CountOccurrences(page.Content, "/F2 8.203 Tf") >= 2,
            "Word-compatible grouped comment balloons should keep the title on the dedicated balloon-label resource (same label typeface, fuller subset) and body text on the regular body resource.");
        TestAssert.Contains("0.973 0.863 0.867 rg", page.Content);
        TestAssert.Contains("0.82 0.204 0.22 RG", page.Content);
        TestAssert.True(
            !page.Content.Contains("0.851 0.592 0 RG", StringComparison.Ordinal),
            "Word-compatible grouped comment balloons should use the Office review palette instead of the author-bucket outline color.");
        TestAssert.True(
            !page.Content.Contains(" 5.5 Tf", StringComparison.Ordinal),
            "Word-compatible grouped comment balloons should not use the legacy summary title font size.");
    }

    public static void DocxWordCompatibleAllMarkupRendersThreadedCommentSeparators()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Threaded comment anchor", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 0, RunChildIndex: 0, TextOffsetInRun: 9, DisplayText: null)
            ]
        };
        DocxRelatedStory parentComment = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Parent threaded comment body", 10d, 12d))],
            [],
            [], null)
        {
            CommentMetadata = new DocxCommentMetadata("Reviewer One", "RO", "2024-01-02T03:04:05Z", "11111111", null, null, true)
        };
        DocxRelatedStory firstReply = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "2",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("First reply body", 10d, 12d))],
            [],
            [], null)
        {
            CommentMetadata = new DocxCommentMetadata("Reviewer Two", "RT", "2024-01-03T03:04:05Z", "22222222", "11111111", "1", false)
        };
        DocxRelatedStory secondReply = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "3",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Second reply body", 10d, 12d))],
            [],
            [], null)
        {
            CommentMetadata = new DocxCommentMetadata("Reviewer Three", "R3", "2024-01-04T03:04:05Z", "33333333", "11111111", "1", false)
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
            RelatedStories = [parentComment, firstReply, secondReply],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxMarkupBalloonPlacementSnapshot placement = renderer.InspectMarkupBalloons(document)
            .Single(item => item.Kind == "Comment");
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        int renderedSeparatorLines = Regex.Matches(page.Content, @"(?<x1>-?[0-9.]+) (?<y>-?[0-9.]+) m (?<x2>-?[0-9.]+) \k<y> l S")
            .Cast<Match>()
            .Count(match =>
            {
                double x1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture);
                double x2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture);
                double y = double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture);
                return x1 >= placement.X &&
                    x2 <= placement.X + placement.Width &&
                    x2 - x1 > placement.Width - 8d &&
                    y > placement.Y &&
                    y < placement.Y + placement.Height;
            });

        TestAssert.Equal(2, placement.CommentReplyCount);
        TestAssert.Equal(2, placement.CommentSeparatorLineCount);
        TestAssert.True(
            placement.Height > 36d && placement.Height < 38d,
            string.Create(CultureInfo.InvariantCulture, $"Threaded comment balloons should expand within the capped Word-compatible height band. Height={placement.Height}."));
        TestAssert.True(
            renderedSeparatorLines >= placement.CommentSeparatorLineCount,
            "Word-compatible threaded comment balloons should emit internal separator strokes for visible reply boundaries.");
    }

    public static void DocxWordCompatibleAllMarkupAnchorsCommentConnectorsAtRangeEnd()
    {
        const string beforeText = "Before anchor ";
        const string rangeText = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega";
        DocxTextRun beforeRun = new(beforeText, 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        DocxTextRun rangeRun = new(rangeText, 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 1,
            SourceTextOffsetInRun = 0
        };
        DocxParagraph paragraph = new(
            [beforeRun, rangeRun],
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
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 2, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ],
            CommentRanges =
            [
                new DocxCommentRange("1", 1, 0, 1, rangeText.Length, 2, 0)
            ],
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "Reviewer", "2026-06-01T00:00:00Z", "inserted", null, [])
            ]
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public comment body", 10d, 12d))],
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
        DocxTextEmissionLineSnapshot[] commentLines = renderer.InspectTextEmission(document).Lines
            .Where(item => item.CommentReferenceCount == 1 && item.Segments.Any(segment => !segment.IsTerminalLineSpace))
            .ToArray();
        TestAssert.True(commentLines.Length >= 2, "The connector fixture should wrap the comment range across multiple emitted lines.");
        DocxTextEmissionSegmentSnapshot firstVisibleSegment = commentLines[0].Segments.First(segment => !segment.IsTerminalLineSpace);
        DocxTextEmissionSegmentSnapshot[] visibleSegments = commentLines[^1].Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();
        DocxTextEmissionSegmentSnapshot lastVisibleSegment = visibleSegments[^1];
        double rangeEndX = visibleSegments[^1].X + visibleSegments[^1].AdvanceProfile.PlannedEmittedAdvance;
        double lastLineAnchorDelta = lastVisibleSegment.BaselineY - placement.AnchorY;
        double firstLineAnchorDelta = firstVisibleSegment.BaselineY - placement.AnchorY;

        TestAssert.Equal(1, placement.CandidateCount);
        // Office: balloon titles land on the anchor row (title ~= row baseline - 0.5),
        // so resolved anchors sit just below the last range line.
        TestAssert.True(
            lastLineAnchorDelta > -2d && lastLineAnchorDelta < 6d && firstLineAnchorDelta > lastLineAnchorDelta + 8d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Word-compatible all-markup should anchor comment connector Y to the wrapped range end line. AnchorY={placement.AnchorY}, FirstBaselineY={firstVisibleSegment.BaselineY}, LastBaselineY={lastVisibleSegment.BaselineY}."));
        // The drawn gap is the 3.18 connector inset plus line-length-dependent emission extras
        // (0.071 positioning spacing accumulated over the wrapped last line plus kern/rounding:
        // observed 7.217 = 3.18 + 4.037 on the 56-gap last line). Queued: resolve anchors from
        // emission space so connectors track the drawn range end exactly.
        TestAssert.True(
            placement.AnchorConnectorX < rangeEndX - 3d &&
            placement.AnchorConnectorX > rangeEndX - 8d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Word-compatible all-markup should anchor comment connectors near the emitted comment range end after connector inset. AnchorX={placement.AnchorConnectorX}, RangeEndX={rangeEndX}."));
    }

    public static void DocxWordCompatibleAllMarkupSuppressesCommentReferenceSpacerTextOperation()
    {
        DocxParagraph paragraph = new(
            [
                new DocxTextRun("Review note", 10d, null, false, false, false, null, null),
                new DocxTextRun(" for measurement.", 10d, null, false, false, false, null, null)
            ],
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
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 0, RunChildIndex: 0, TextOffsetInRun: 11, DisplayText: null)
            ]
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public comment body", 10d, 12d))],
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
            [new DocxParagraphElement(paragraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxTextEmissionSnapshot preserve = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectTextEmission(document);
        DocxTextEmissionSnapshot wordCompatible = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectTextEmission(document);

        TestAssert.True(
            DocxTests.HasCommentReferenceSpacerTextOperation(preserve),
            "The fixture should expose a standalone comment-reference spacer before Word-compatible filtering.");
        TestAssert.True(
            !DocxTests.HasCommentReferenceSpacerTextOperation(wordCompatible),
            "Word-compatible all-markup should keep the visual gap without emitting a separate spacer text operation.");
        double printScale = DocxRenderer.ResolveWordCompatiblePrintScale(document, DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
        TestAssert.True(
            Math.Abs(DocxTests.FirstTextEmissionX(wordCompatible) - DocxTests.FirstTextEmissionX(preserve) * printScale) < 0.5d,
            "Word-compatible all-markup should fit body text X by the print scale (Office: left margin times scale).");
    }

    public static void DocxAllMarkupRendererDrawsRevisionBalloons()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot(document.Paragraphs.Single().Revisions), "RG"), page.Content);
        TestAssert.Contains(" re B*", page.Content);
        TestAssert.Contains(" l S", page.Content);
        TestAssert.Contains("Inserted text", DocxRenderer.BuildRevisionBalloonPreview(document.Paragraphs.Single().Revisions));
        TestAssert.Contains("Deleted text", DocxRenderer.BuildRevisionBalloonPreview(document.Paragraphs.Single().Revisions));
        TestAssert.Contains("Moved from", DocxRenderer.BuildRevisionBalloonPreview(document.Paragraphs.Single().Revisions));
        TestAssert.Contains("Moved to", DocxRenderer.BuildRevisionBalloonPreview(document.Paragraphs.Single().Revisions));
        TestAssert.Contains("4 reviewers", DocxRenderer.BuildRevisionBalloonTitle(document.Paragraphs.Single().Revisions));
        TestAssert.Contains("Deleted: \"Deleted\"", DocxRenderer.BuildRevisionBalloonPreview(document.Paragraphs.Single()));
        TestAssert.Contains("Moved from: \"MovedFrom\"", DocxRenderer.BuildRevisionBalloonPreview(document.Paragraphs.Single()));
    }

    public static void DocxAllMarkupRendererPlacesMoveRevisionBalloonsWithConnectors()
    {
        var moveFrom = new DocxRevisionInfo(DocxRevisionKind.MoveFrom, "701", "Reviewer", "2026-06-10T00:00:00Z", "moveFrom", null, []);
        var moveTo = new DocxRevisionInfo(DocxRevisionKind.MoveTo, "702", "Reviewer", "2026-06-10T00:00:00Z", "moveTo", null, []);
        DocxParagraph movedFrom = CreateMovedParagraph("Moved from text", moveFrom);
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Neutral paragraph", 10d, 12d);
        DocxParagraph movedTo = CreateMovedParagraph("Moved to text", moveTo);
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
            [
                new DocxParagraphElement(movedFrom),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(movedTo)
            ],
            [movedFrom, filler, movedTo],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] revisionPlacements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.RevisionCandidateCount != 0)
            .ToArray();

        TestAssert.True(revisionPlacements.Length >= 1, "Moved-from and moved-to revisions should produce markup balloon placements.");
        TestAssert.Equal(2, revisionPlacements.Sum(placement => placement.RevisionCandidateCount));
        TestAssert.True(
            revisionPlacements.All(placement =>
                placement.AnchorConnectorX > 0d &&
                placement.BalloonConnectorX > 0d &&
                placement.AnchorY > 0d &&
                placement.Width > 0d &&
                placement.Height > 0d),
            "Moved-from and moved-to revision balloons should expose concrete connector geometry.");
        TestAssert.Contains("Moved from: \"Moved from text\"", DocxRenderer.BuildRevisionBalloonPreview(movedFrom));
        TestAssert.Contains("Moved to: \"Moved to text\"", DocxRenderer.BuildRevisionBalloonPreview(movedTo));

        static DocxParagraph CreateMovedParagraph(string text, DocxRevisionInfo revision)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d) with
            {
                Runs =
                [
                    new DocxTextRun(text, 10d, null, false, false, false, null, null)
                    {
                        Revision = revision
                    }
                ],
                Revisions = [revision]
            };
        }
    }

    public static void DocxAllMarkupRevisionBalloonPreviewUsesAggregateRunProvenance()
    {
        var insertion = new DocxRevisionInfo(DocxRevisionKind.Insertion, "4", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        var deletion = new DocxRevisionInfo(DocxRevisionKind.Deletion, "1", "Reviewer", "2026-06-10T00:00:00Z", "del", null, []);
        var moveFrom = new DocxRevisionInfo(DocxRevisionKind.MoveFrom, "2", "Reviewer", "2026-06-10T00:00:00Z", "moveFrom", null, []);
        var moveTo = new DocxRevisionInfo(DocxRevisionKind.MoveTo, "3", "Reviewer", "2026-06-10T00:00:00Z", "moveTo", null, []);
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Deleted MovedFrom MovedTo Inserted", 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun("Deleted", 10d, null, false, false, false, null, null)
                {
                    Revisions = [deletion]
                },
                new DocxTextRun(" ", 10d, null, false, false, false, null, null),
                new DocxTextRun("MovedFrom", 10d, null, false, false, false, null, null)
                {
                    Revisions = [moveFrom]
                },
                new DocxTextRun(" ", 10d, null, false, false, false, null, null),
                new DocxTextRun("MovedTo", 10d, null, false, false, false, null, null)
                {
                    Revisions = [moveTo]
                },
                new DocxTextRun(" ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Inserted", 10d, null, false, false, false, null, null)
                {
                    Revisions = [insertion]
                }
            ],
            Revisions = [moveTo, deletion, insertion, moveFrom]
        };

        string preview = DocxRenderer.BuildRevisionBalloonPreview(paragraph);

        TestAssert.Equal(
            "Inserted text, Deleted text, Moved from, Moved to, Deleted: \"Deleted\", Moved from: \"MovedFrom\", Moved to: \"MovedTo\"",
            preview);
    }

    public static void DocxAllMarkupRevisionBalloonPreviewSummarizesLongDeletedText()
    {
        var deletion = new DocxRevisionInfo(DocxRevisionKind.Deletion, "1", "Reviewer", "2026-06-10T00:00:00Z", "del", null, []);
        const string longDeletedText = "Deleted content that is intentionally too long for an inline balloon quote";
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph(longDeletedText, 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun(longDeletedText, 10d, null, false, false, false, null, null)
                {
                    Revisions = [deletion]
                }
            ],
            Revisions = [deletion]
        };

        string preview = DocxRenderer.BuildRevisionBalloonPreview(paragraph);

        TestAssert.Contains("Deleted text", preview);
        TestAssert.DoesNotContain("Deleted: \"", preview);
        TestAssert.DoesNotContain(longDeletedText, preview);
    }

    public static void DocxMarkupRevisionColorsUseStableAuthorBuckets()
    {
        (byte Red, byte Green, byte Blue) first = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
        [
            new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "Reviewer A", "2026-06-10T00:00:00Z", "ins", null, [])
        ]);
        (byte Red, byte Green, byte Blue) firstAgain = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
        [
            new DocxRevisionInfo(DocxRevisionKind.Deletion, "2", "Reviewer A", "2026-06-11T00:00:00Z", "del", null, [])
        ]);
        (byte Red, byte Green, byte Blue) second = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
        [
            new DocxRevisionInfo(DocxRevisionKind.Insertion, "3", "Reviewer B", "2026-06-12T00:00:00Z", "ins", null, [])
        ]);
        (byte Red, byte Green, byte Blue) firstWithEquivalentWhitespace = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
        [
            new DocxRevisionInfo(DocxRevisionKind.Deletion, "4", "  reviewer\t  a  ", "2026-06-12T00:00:00Z", "del", null, [])
        ]);
        (byte Red, byte Green, byte Blue) composedUnicode = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
        [
            new DocxRevisionInfo(DocxRevisionKind.Insertion, "5", "JOS\u00c9", "2026-06-12T00:00:00Z", "ins", null, [])
        ]);
        (byte Red, byte Green, byte Blue) decomposedUnicode = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
        [
            new DocxRevisionInfo(DocxRevisionKind.Deletion, "6", "Jose\u0301", "2026-06-12T00:00:00Z", "del", null, [])
        ]);
        (byte Red, byte Green, byte Blue) fallback = DocxRenderer.ResolveRevisionAuthorColorSnapshot([]);

        TestAssert.Equal(first, firstAgain);
        TestAssert.Equal(first, firstWithEquivalentWhitespace);
        TestAssert.Equal(composedUnicode, decomposedUnicode);
        TestAssert.True(!first.Equals(second), "Different reviewer buckets should usually receive different revision colors in the built-in palette.");
        TestAssert.True(fallback.Red != 0 || fallback.Green != 0 || fallback.Blue != 0, "Authorless revisions should still receive a visible deterministic color.");
    }

    public static void DocxMarkupRevisionColorsUseDominantAuthorBucketIndependentOfOrder()
    {
        var reviewerAInsertion = new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "Reviewer A", "2026-06-10T00:00:00Z", "ins", null, []);
        var reviewerADeletion = new DocxRevisionInfo(DocxRevisionKind.Deletion, "2", "Reviewer A", "2026-06-11T00:00:00Z", "del", null, []);
        var reviewerBInsertion = new DocxRevisionInfo(DocxRevisionKind.Insertion, "3", "Reviewer B", "2026-06-12T00:00:00Z", "ins", null, []);
        (byte Red, byte Green, byte Blue) reviewerA = DocxRenderer.ResolveRevisionAuthorColorSnapshot([reviewerAInsertion]);
        (byte Red, byte Green, byte Blue) dominantA = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
            [reviewerBInsertion, reviewerAInsertion, reviewerADeletion]);
        (byte Red, byte Green, byte Blue) dominantAReordered = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
            [reviewerADeletion, reviewerBInsertion, reviewerAInsertion]);
        (byte Red, byte Green, byte Blue) tiedAuthors = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
            [reviewerBInsertion, reviewerAInsertion]);
        (byte Red, byte Green, byte Blue) tiedAuthorsReordered = DocxRenderer.ResolveRevisionAuthorColorSnapshot(
            [reviewerAInsertion, reviewerBInsertion]);

        TestAssert.Equal(reviewerA, dominantA);
        TestAssert.Equal(dominantA, dominantAReordered);
        TestAssert.Equal(tiedAuthors, tiedAuthorsReordered);
        TestAssert.Equal(reviewerA, tiedAuthors);
    }

    public static void DocxMarkupBalloonLayoutSelectsSideStacksAnchorsAndHandlesOverflow()
    {
        var bodyElements = new List<DocxBodyElement>();
        var relatedStories = new List<DocxRelatedStory>();
        for (int i = 0; i < 12; i++)
        {
            string id = (i + 1).ToString(CultureInfo.InvariantCulture);
            DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Anchor " + id, 10d, 12d) with
            {
                InlineReferences =
                [
                    new DocxInlineReference(DocxRelatedStoryKind.Comment, id, null, SourceRunIndex: 0, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
                ],
                Revisions =
                [
                    new DocxRevisionInfo(DocxRevisionKind.Insertion, id, "A", "2026-06-10T00:00:00Z", "ins", null, [])
                ]
            };
            bodyElements.Add(new DocxParagraphElement(paragraph));

            DocxParagraph commentParagraph = DocxTests.CreateDocxLayoutParagraph("Comment body " + id, 10d, 12d);
            relatedStories.Add(new DocxRelatedStory(
                DocxRelatedStoryKind.Comment,
                "/word/comments.xml",
                id,
                [new DocxParagraphElement(commentParagraph)],
                [],
                [], null));
        }

        var document = new DocxDocument(
            200d,
            185d,
            50d,
            15d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            bodyElements,
            [],
            [])
        {
            RelatedStories = relatedStories,
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .ToArray();
        DocxMarkupBalloonPlacementSnapshot[] balloons = placements
            .Where(placement => !placement.IsOverflowSummary)
            .ToArray();

        TestAssert.True(balloons.Length > 1, "The fixture should place multiple markup balloons before overflowing.");
        TestAssert.True(placements.Any(placement => placement.IsOverflowSummary), "Overflowing markup should be summarized instead of overlapping at the margin bottom.");
        TestAssert.True(placements.Length < 24, "Overflow handling should suppress balloons that do not fit on the page.");
        TestAssert.True(placements.Where(placement => placement.IsOverflowSummary).All(placement => placement.CandidateCount > 0 && placement.OverflowStartIndex is not null && placement.OverflowEndIndex is not null), "Overflow summaries should expose private-safe continuation ranges.");
        TestAssert.Equal(24, placements.Sum(placement => placement.CandidateCount));
        DocxMarkupBalloonPlacementSnapshot[] overflowSummaries = placements
            .Where(placement => placement.IsOverflowSummary)
            .OrderBy(placement => placement.OverflowStartIndex)
            .ToArray();
        int nextOverflowStartIndex = 1;
        foreach (DocxMarkupBalloonPlacementSnapshot overflowSummary in overflowSummaries)
        {
            TestAssert.Equal(nextOverflowStartIndex, overflowSummary.OverflowStartIndex ?? 0);
            TestAssert.True(
                overflowSummary.OverflowEndIndex >= overflowSummary.OverflowStartIndex,
                "Overflow continuation ranges should be non-empty.");
            nextOverflowStartIndex = (overflowSummary.OverflowEndIndex ?? nextOverflowStartIndex) + 1;
        }

        TestAssert.True(placements.All(placement => placement.Side == "Left"), "A wider left margin should be selected for markup balloons.");
        TestAssert.True(placements.All(placement => placement.X >= 0d && placement.X + placement.Width <= document.MarginLeftPoints + 0.001d), "Left-side balloons should stay inside the available margin area.");
        TestAssert.True(
            balloons.Any(placement => placement.Kind == "Markup" && placement.CandidateCount > 1),
            "Same-anchor comment/revision markup should be grouped into a single private-safe markup balloon.");

        for (int i = 0; i < balloons.Length - 1; i++)
        {
            TestAssert.True(balloons[i].AnchorY >= balloons[i + 1].AnchorY, "Balloon placements should stay ordered by anchor position from top to bottom.");
        }

        DocxMarkupBalloonPlacementSnapshot[] byVerticalPosition = balloons
            .OrderByDescending(placement => placement.Y)
            .ToArray();
        for (int i = 0; i < byVerticalPosition.Length - 1; i++)
        {
            TestAssert.True(
                byVerticalPosition[i].Y >= byVerticalPosition[i + 1].Y + byVerticalPosition[i + 1].Height,
                "Balloon placements should not collide vertically.");
        }
    }

    public static void DocxMarkupBalloonLayoutGroupsNearbyRevisionAnchors()
    {
        var bodyElements = new List<DocxBodyElement>();
        for (int i = 0; i < 4; i++)
        {
            string id = (i + 1).ToString(CultureInfo.InvariantCulture);
            DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Dense " + id, 5d, 5d) with
            {
                Revisions =
                [
                    new DocxRevisionInfo(DocxRevisionKind.Insertion, id, "A", "2026-06-10T00:00:00Z", "ins", null, [])
                ]
            };
            bodyElements.Add(new DocxParagraphElement(paragraph));
        }

        var document = new DocxDocument(
            220d,
            240d,
            55d,
            15d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            bodyElements,
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] revisionBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.Kind == "Revision")
            .ToArray();

        TestAssert.True(revisionBalloons.Any(placement => placement.CandidateCount > 1), "Nearby dense revision anchors should be grouped into a single private-safe markup balloon.");
        TestAssert.True(revisionBalloons.Length < 4, "Grouping should reduce one-balloon-per-revision churn in dense review pages.");
        TestAssert.True(revisionBalloons.Any(placement => placement.CandidateCount == 3), "Nearby dense revision anchors should group into capped chunks before starting a new balloon.");
        TestAssert.True(revisionBalloons.All(placement => placement.CandidateCount <= 3), "Nearby revision grouping should cap dense chunks instead of hiding too many edits behind one balloon.");
    }

    public static void DocxWordCompatibleAllMarkupSkipsInlineRevisionBalloons()
    {
        // Office A/B (dense-revisions, balloon-lane-bands, review references): Word shows
        // insertions, deletions, and moves inline with zero revision balloons.
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Dense ins del move", 5d, 5d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "A", "2026-06-10T00:00:00Z", "ins", null, []),
                new DocxRevisionInfo(DocxRevisionKind.Deletion, "2", "A", "2026-06-10T00:00:00Z", "del", null, []),
                new DocxRevisionInfo(DocxRevisionKind.MoveFrom, "3", "A", "2026-06-10T00:00:00Z", "moveFrom", null, [])
            ]
        };
        var document = new DocxDocument(
            220d,
            240d,
            55d,
            15d,
            10d,
            10d,
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

        DocxMarkupBalloonPlacementSnapshot[] revisionBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.Kind == "Revision")
            .ToArray();

        TestAssert.Equal(0, revisionBalloons.Length);
    }

    public static void DocxWordCompatibleAllMarkupSkipsVoidPropertyChangeBalloons()
    {
        // Office A/B (review reference): an empty rPrChange sustains no Word balloon.
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Void format anchor", 5d, 5d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "A", "2026-06-10T00:00:00Z", "ins", null, []),
                new DocxRevisionInfo(DocxRevisionKind.RunPropertiesChange, "2", "A", "2026-06-10T00:00:00Z", "rPrChange", null, [])
            ]
        };
        var document = new DocxDocument(
            220d,
            240d,
            55d,
            15d,
            10d,
            10d,
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

        DocxMarkupBalloonPlacementSnapshot[] revisionBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.Kind == "Revision")
            .ToArray();

        TestAssert.Equal(0, revisionBalloons.Length);
    }

    public static void DocxBalloonTextResourceCoversSyntheticTitleGlyphs()
    {
        // Office A/B: balloon titles are synthetic (Commented [RV1]: ) and must render fully
        // even when no body run contains their glyphs.
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("aaa", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 0, RunChildIndex: 0, TextOffsetInRun: 1, DisplayText: null)
            ]
        };
        DocxParagraph commentParagraph = DocxTests.CreateDocxLayoutParagraph("bbb", 10d, 12d);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(commentParagraph)],
            [],
            [], null)
        {
            CommentMetadata = new DocxCommentMetadata("Reviewer", "RV", "2026-06-01T00:00:00Z", null, null, null, null)
        };
        var document = new DocxDocument(
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
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        const string title = "Commented [RV1]: ";
        TestAssert.True(page.Fonts.Any(resource => CoversAllTitleGlyphs(resource.Font, title)),
            "Some page font must cover every synthetic balloon title glyph.");
    }

    private static bool CoversAllTitleGlyphs(PdfEmbeddedFont font, string text)
    {
        foreach (System.Text.Rune rune in text.EnumerateRunes())
        {
            ushort glyph = font.Font.MapCodePoint(rune.Value);
            if (glyph == 0 || !font.TryGetEncodedCid(glyph, out _))
            {
                return false;
            }
        }

        return true;
    }

    public static void DocxWordCompatibleAllMarkupKeepsFormattingRevisionBalloons()
    {
        // Office A/B (grounded review reference): Word balloons formatting revisions.
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Formatted run", 5d, 5d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.RunPropertiesChange, "9", "A", "2026-06-10T00:00:00Z", "rPrChange", null, propertyElementNames: ["b"])
            ]
        };
        var document = new DocxDocument(
            220d,
            240d,
            55d,
            15d,
            10d,
            10d,
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

        DocxMarkupBalloonPlacementSnapshot[] revisionBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.Kind == "Revision")
            .ToArray();

        TestAssert.Equal(1, revisionBalloons.Length);
    }

    public static void DocxMarkupBalloonLayoutOffsetsNearbyMixedConnectorAnchors()
    {
        DocxParagraph commentParagraph = DocxTests.CreateDocxLayoutParagraph("Shared width", 10d, 8d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: -1, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ]
        };
        DocxParagraph revisionParagraph = DocxTests.CreateDocxLayoutParagraph("Shared width", 10d, 8d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "Reviewer", "2026-06-01T00:00:00Z", "inserted", null, [])
            ]
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public collision comment", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            260d,
            240d,
            30d,
            90d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(commentParagraph), new DocxParagraphElement(revisionParagraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .Where(placement => !placement.IsOverflowSummary)
            .OrderByDescending(placement => placement.AnchorY)
            .ToArray();

        TestAssert.Equal(2, placements.Length);
        TestAssert.True(
            placements.Any(placement => placement.CommentCandidateCount == 1) &&
            placements.Any(placement => placement.RevisionCandidateCount == 1),
            "The collision fixture should keep separate nearby comment and revision balloons.");
        TestAssert.True(
            Math.Abs(placements[0].AnchorY - placements[1].AnchorY) < 9d,
            "The collision fixture should place mixed connector anchors inside the near-anchor collision band.");
        TestAssert.True(
            Math.Abs(placements[0].AnchorConnectorX - placements[1].AnchorConnectorX) >= 1d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Nearby mixed comment/revision connector anchors should be offset horizontally. FirstX={placements[0].AnchorConnectorX}, SecondX={placements[1].AnchorConnectorX}."));
    }

    public static void DocxMarkupBalloonLayoutPrioritizesCommentsOverNearbyRevisions()
    {
        DocxParagraph revisionParagraph = DocxTests.CreateDocxLayoutParagraph("Priority revision", 10d, 8d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "Reviewer", "2026-06-01T00:00:00Z", "inserted", null, [])
            ]
        };
        DocxParagraph commentParagraph = DocxTests.CreateDocxLayoutParagraph("Priority comment", 10d, 8d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: -1, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ]
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public priority comment", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            260d,
            240d,
            30d,
            90d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(revisionParagraph), new DocxParagraphElement(commentParagraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .Where(placement => !placement.IsOverflowSummary)
            .ToArray();

        TestAssert.Equal(2, placements.Length);
        TestAssert.True(
            placements[1].AnchorY > placements[0].AnchorY &&
            placements[1].AnchorY - placements[0].AnchorY < 9d,
            "The fixture should put the revision anchor slightly above the comment anchor inside one competition band.");
        TestAssert.Equal("Comment", placements[0].Kind);
        TestAssert.Equal("Revision", placements[1].Kind);
    }

    public static void DocxMarkupBalloonLayoutSeparatesDistantAnchorsIntoLaneBands()
    {
        DocxParagraph upperParagraph = DocxTests.CreateDocxLayoutParagraph("Upper revision", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "upper", "Reviewer", "2026-06-01T00:00:00Z", "inserted", null, [])
            ]
        };
        DocxParagraph spacerParagraph = DocxTests.CreateDocxLayoutParagraph("Spacer", 1d, 80d);
        DocxParagraph lowerParagraph = DocxTests.CreateDocxLayoutParagraph("Lower revision", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Deletion, "lower", "Reviewer", "2026-06-01T00:00:00Z", "deleted", null, [])
            ]
        };
        DocxDocument document = new(
            260d,
            300d,
            30d,
            90d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(upperParagraph),
                new DocxParagraphElement(spacerParagraph),
                new DocxParagraphElement(lowerParagraph)
            ],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .Where(placement => !placement.IsOverflowSummary)
            .OrderBy(placement => placement.LaneBandIndex)
            .ToArray();

        TestAssert.Equal(2, placements.Length);
        TestAssert.Equal(0, placements[0].LaneBandIndex);
        TestAssert.Equal(1, placements[1].LaneBandIndex);
        TestAssert.True(
            placements.All(placement => placement.LaneBandCandidateCount == placement.CandidateCount),
            "Separated lane bands should expose private-safe candidate counts independently.");
        TestAssert.True(
            placements[0].Y >= placements[1].Y + placements[1].Height + 18d,
            "Distant review-anchor lane bands should remain visually separated instead of being treated as one collision band.");
    }
    public static void DocxWordCompatibleAllMarkupSkipsFloatingTextBoxCommentBalloons()
    {
        // Office A/B (w6-tbxctl probe, Word-COM rendered): Word balloons body-anchored
        // comments but never body-flow floating-textbox ones, so a textbox-only comment
        // sustains no Word-compatible balloon and reserves no lane.
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateCommentMarkerParagraph("Text box review anchor", "5");
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "5",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Text box comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] commentBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.Kind == "Comment")
            .ToArray();

        TestAssert.Equal(0, commentBalloons.Length);
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(document, context));
    }

    public static void DocxWordCompatibleAllMarkupKeepsBodyCommentBalloonAlongsideFloatingTextBoxComment()
    {
        // Companion to the suppression test: a body-anchored comment still balloons and
        // reserves the lane when a floating-textbox comment is present.
        DocxParagraph bodyParagraph = DocxTests.CreateCommentMarkerParagraph("Body review anchor", "6");
        DocxParagraph textBoxParagraph = DocxTests.CreateCommentMarkerParagraph("Text box review anchor", "5");
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxRelatedStory bodyCommentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "6",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body comment body", 10d, 12d))],
            [],
            [], null);
        DocxRelatedStory textBoxCommentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "5",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Text box comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            RelatedStories = [bodyCommentStory, textBoxCommentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] commentBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.Kind == "Comment")
            .ToArray();

        TestAssert.Equal(1, commentBalloons.Length);
        TestAssert.Equal(1, commentBalloons.Single().CandidateCount);
        TestAssert.Equal(1, commentBalloons.Single().CommentCandidateCount);
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.True(DocxRenderer.ResolveWordCompatiblePrintScale(document, context) < 1d, "A body-anchored comment should reserve the balloon lane.");
    }
    public static void DocxWordCompatibleAllMarkupSkipsStaticFloatingTextBoxCommentBalloons()
    {
        // Office A/B (w6-staticfloat probe, Word-COM rendered): like body-flow floating
        // textboxes, static ones sustain no Word-compatible balloons or lane.
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateCommentMarkerParagraph("Static text box review anchor", "7");
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)])]
            }
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "7",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Static text box comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            pageSettings,
            [],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxMarkupBalloonPlacementSnapshot[] commentBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectMarkupBalloons(document)
            .Where(placement => placement.Kind == "Comment")
            .ToArray();

        TestAssert.Equal(0, commentBalloons.Length);
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(document, context));
    }
}
