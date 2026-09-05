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

internal static class DocxHyperlinksTests
{
    public static void DocxRendererEmitsBodyExternalHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Before ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null),
            new DocxTextRun(" After", 10d, null, false, false, false, null, null)
        };
        var paragraph = new DocxParagraph(
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
                new DocxHyperlinkSpan("rIdLink", null, null, null, "https://example.invalid/docx", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/docx", annotation.Uri);
        TestAssert.True(annotation.X > document.MarginLeftPoints, "The annotation should be anchored to the placed hyperlink run, not the paragraph origin.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover the rendered hyperlink text.");
        TestAssert.True(annotation.Height > 0d, "The annotation should use font metrics for a non-empty rectangle.");
    }

    public static void DocxRendererUsesSourceRunIndexesForHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Before ", 10d, null, false, false, false, null, null)
            {
                SourceRunIndex = 10
            },
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
            {
                SourceRunIndex = 42
            }
        };
        var paragraph = new DocxParagraph(
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
                new DocxHyperlinkSpan("rIdLink", null, null, null, "https://example.invalid/source-run", "External", null, 42, 1, 1, 1, 4)
            ]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/source-run", annotation.Uri);
        TestAssert.True(annotation.X > document.MarginLeftPoints, "Hyperlink annotations should match rendered source-run indexes after filtering changes text-run positions.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover the rendered source-run hyperlink text.");
    }

    public static void DocxWordCompatibleAllMarkupUsesEmittedAdvanceForHyperlinkAnnotations()
    {
        const string linkText = "LinkedWords";
        var runs = new[]
        {
            new DocxTextRun("Before ", 10d, null, false, false, false, null, null),
            new DocxTextRun(linkText, 10d, null, false, false, false, null, null)
        };
        var paragraph = new DocxParagraph(
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
                new DocxHyperlinkSpan("rIdLink", null, null, null, "https://example.invalid/docx", "External", null, 1, 1, 1, 1, linkText.Length)
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

        DocxTextEmissionSegmentSnapshot linkSegment = renderer
            .InspectTextEmission(document)
            .Lines
            .SelectMany(line => line.Segments)
            .Single(segment => !segment.IsTerminalLineSpace && segment.TextLength == linkText.Length);
        PdfLinkAnnotation annotation = renderer.RenderBlankPages(document, null, CancellationToken.None).Single().Annotations.Single();

        TestAssert.True(
            Math.Abs(annotation.Width - linkSegment.AdvanceProfile.PlannedEmittedAdvance) < 0.001d,
            "Word-compatible all-markup hyperlink annotations should cover the emitted glyph advance after positioned spacing.");
        TestAssert.True(
            Math.Abs(annotation.Width - linkSegment.Width) > 0.05d,
            "The regression should exercise a link whose emitted advance differs from the layout segment width.");
    }

    public static void DocxWordCompatibleAllMarkupUsesEmittedAdvanceForFloatingTextBoxHyperlinkAnnotations()
    {
        const string linkText = "LinkedWords";
        DocxParagraph textBoxParagraph = new(
            [
                new DocxTextRun("Before ", 10d, null, false, false, false, null, null),
                new DocxTextRun(linkText, 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdTextBoxLink", null, null, null, "https://example.invalid/textbox-word-compatible", "External", null, 1, 1, 1, 1, linkText.Length)
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
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxTextEmissionSegmentSnapshot linkSegment = renderer
            .InspectTextEmission(document)
            .Lines
            .Single(line => line.StoryKind == "TextBox" && line.ContainerStoryKind == "Body")
            .Segments
            .Single(segment => !segment.IsTerminalLineSpace && segment.TextLength == linkText.Length);
        PdfLinkAnnotation annotation = renderer.RenderBlankPages(document, null, CancellationToken.None).Single().Annotations.Single();

        TestAssert.True(
            Math.Abs(annotation.X - linkSegment.X) < 0.001d,
            "Word-compatible all-markup floating text-box hyperlink annotations should use emitted text-box segment x coordinates.");
        TestAssert.True(
            Math.Abs(annotation.Width - linkSegment.AdvanceProfile.PlannedEmittedAdvance) < 0.001d,
            "Word-compatible all-markup floating text-box hyperlink annotations should cover the emitted glyph advance after positioned spacing.");
        TestAssert.True(
            Math.Abs(annotation.Width - linkSegment.Width) > 0.05d,
            "The regression should exercise a floating text-box link whose emitted advance differs from the layout segment width.");
    }

    public static void DocxWordCompatibleAllMarkupUsesEmittedAdvanceForStaticTextBoxHyperlinkAnnotations()
    {
        const string linkText = "LinkedWords";
        DocxParagraph textBoxParagraph = new(
            [
                new DocxTextRun("Header ", 10d, null, false, false, false, null, null),
                new DocxTextRun(linkText, 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan("rIdStaticTextBoxLink", null, null, null, "https://example.invalid/static-textbox-word-compatible", "External", null, 1, 1, 1, 1, linkText.Length)
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
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxTextEmissionSegmentSnapshot linkSegment = renderer
            .InspectTextEmission(document)
            .Lines
            .Single(line => line.IsStaticStory && line.StoryKind == "TextBox" && line.ContainerStoryKind == "Header")
            .Segments
            .Single(segment => !segment.IsTerminalLineSpace && segment.TextLength == linkText.Length);
        PdfLinkAnnotation annotation = renderer.RenderBlankPages(document, null, CancellationToken.None).Single().Annotations.Single();

        TestAssert.True(
            Math.Abs(annotation.X - linkSegment.X) < 0.001d,
            "Word-compatible all-markup static text-box hyperlink annotations should use emitted text-box segment x coordinates.");
        TestAssert.True(
            Math.Abs(annotation.Width - linkSegment.AdvanceProfile.PlannedEmittedAdvance) < 0.001d,
            "Word-compatible all-markup static text-box hyperlink annotations should cover the emitted glyph advance after positioned spacing.");
        TestAssert.True(
            Math.Abs(annotation.Width - linkSegment.Width) > 0.05d,
            "The regression should exercise a static text-box link whose emitted advance differs from the layout segment width.");
    }

    public static void DocxRendererDoesNotEmitUriAnnotationsForInternalHyperlinks()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Internal", 10d, 12d) with
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "Bookmark", null, null, null, null, null, 0, 1, 0, 1, 8)
            ]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(0, page.Annotations.Count);
    }

    public static void DocxRendererEmitsInternalHyperlinkDestinationsFromBookmarks()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Target", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "BookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 6)
            ]
        };
        DocxParagraph targetParagraph = DocxTests.CreateDocxLayoutParagraph("Target text", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("3", "BookmarkTarget", 0, 0, 0)
            ]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(linkParagraph), new DocxParagraphElement(targetParagraph)],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal DOCX links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal DOCX links should resolve to a PDF page destination.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed bookmark text coordinates.");
        TestAssert.True(annotation.Destination?.Top > 0d, "The destination should point to a concrete bookmark line top.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should still cover the rendered internal-link text.");
    }

    public static void DocxRendererUsesSourceRunIndexesForBookmarkDestinations()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Target", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "SparseBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 6)
            ]
        };
        var targetRuns = new[]
        {
            new DocxTextRun("Before ", 10d, null, false, false, false, null, null)
            {
                SourceRunIndex = 10
            },
            new DocxTextRun("Target", 10d, null, false, false, false, null, null)
            {
                SourceRunIndex = 42
            }
        };
        var targetParagraph = new DocxParagraph(
            targetRuns,
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
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("4", "SparseBookmarkTarget", 42, 1, 0)
            ]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(linkParagraph), new DocxParagraphElement(targetParagraph)],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal links should resolve to a PDF page destination.");
        TestAssert.True(
            annotation.Destination?.Left > document.MarginLeftPoints,
            "Bookmark destinations should match rendered source-run indexes after filtering changes text-run positions.");
    }

    public static void DocxRendererUsesSourceTextOffsetsForBookmarkDestinations()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Target", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "OffsetBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 6)
            ]
        };
        DocxTextRun targetRun = new("Before Target", 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 42,
            SourceTextOffsetInRun = 0
        };
        var targetParagraph = new DocxParagraph(
            [targetRun],
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
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("5", "OffsetBookmarkTarget", 42, 0, "Before ".Length)
            ]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(linkParagraph), new DocxParagraphElement(targetParagraph)],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal links should resolve to a PDF page destination.");
        TestAssert.True(
            annotation.Destination?.Left > document.MarginLeftPoints,
            "Bookmark destinations should honor the source text offset when the bookmark starts inside a rendered run.");
    }

    public static void DocxRendererEmitsPlacedFootnoteExternalHyperlinkAnnotations()
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
        var footnoteRuns = new[]
        {
            new DocxTextRun("Footnote ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var footnoteParagraph = new DocxParagraph(
            footnoteRuns,
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
                new DocxHyperlinkSpan("rIdFootnoteLink", null, null, null, "https://example.invalid/footnote", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "9",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [footnoteStory]
            };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/footnote", annotation.Uri);
        TestAssert.True(annotation.X >= document.MarginLeftPoints, "The footnote annotation should use placed note x coordinates.");
        TestAssert.True(annotation.Y >= document.MarginBottomPoints, "The footnote annotation should be anchored inside the placed note region.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover placed footnote hyperlink text.");
    }

    public static void DocxRendererEmitsPlacedEndnoteExternalHyperlinkAnnotations()
    {
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
        var endnoteRuns = new[]
        {
            new DocxTextRun("Endnote ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var endnoteParagraph = new DocxParagraph(
            endnoteRuns,
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
                new DocxHyperlinkSpan("rIdEndnoteLink", null, null, null, "https://example.invalid/endnote", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "11",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [endnoteStory]
            };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/endnote", annotation.Uri);
        TestAssert.True(annotation.X >= document.MarginLeftPoints, "The endnote annotation should use placed note x coordinates.");
        TestAssert.True(annotation.Y >= document.MarginBottomPoints, "The endnote annotation should be anchored inside the placed note region.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover placed endnote hyperlink text.");
    }

    public static void DocxRendererEmitsPlacedFootnoteInternalHyperlinkDestinations()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Footnote target", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "FootnoteBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 15)
            ]
        };
        DocxParagraph markerParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
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
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph targetParagraph = DocxTests.CreateDocxLayoutParagraph("Bookmark target", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("12", "FootnoteBookmarkTarget", 0, 0, 0)
            ]
        };
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "12",
            [new DocxParagraphElement(targetParagraph)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(linkParagraph), new DocxParagraphElement(markerParagraph)],
                [])
            with
            {
                RelatedStories = [footnoteStory]
            };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal placed-footnote links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal links should resolve to placed footnote bookmark destinations.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed footnote x coordinates.");
        TestAssert.True(annotation.Destination?.Top >= document.MarginBottomPoints, "The destination should use placed footnote y coordinates.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should cover the rendered body internal-link text.");
    }

    public static void DocxRendererEmitsPlacedEndnoteInternalHyperlinkDestinations()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Endnote target", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "EndnoteBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 14)
            ]
        };
        DocxParagraph markerParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "13",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph targetParagraph = DocxTests.CreateDocxLayoutParagraph("Bookmark target", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("13", "EndnoteBookmarkTarget", 0, 0, 0)
            ]
        };
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "13",
            [new DocxParagraphElement(targetParagraph)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(linkParagraph), new DocxParagraphElement(markerParagraph)],
                [])
            with
            {
                RelatedStories = [endnoteStory]
            };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal placed-endnote links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal links should resolve to placed endnote bookmark destinations.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed endnote x coordinates.");
        TestAssert.True(annotation.Destination?.Top >= document.MarginBottomPoints, "The destination should use placed endnote y coordinates.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should cover the rendered body internal-link text.");
    }

    public static void DocxRendererEmitsFloatingTextBoxExternalHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Before ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var textBoxParagraph = new DocxParagraph(
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
                new DocxHyperlinkSpan("rIdTextBox", null, null, null, "https://example.invalid/textbox", "External", null, 1, 1, 1, 1, 4)
            ]
        };
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

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/textbox", annotation.Uri);
        TestAssert.True(annotation.X > document.MarginLeftPoints, "Floating text-box annotation should be anchored to the placed text-box line, not the page origin.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover floating text-box hyperlink text.");
    }

    public static void DocxRendererEmitsStaticFloatingTextBoxExternalHyperlinkAnnotations()
    {
        var runs = new[]
        {
            new DocxTextRun("Header ", 10d, null, false, false, false, null, null),
            new DocxTextRun("Link", 10d, null, false, false, false, null, null)
        };
        var textBoxParagraph = new DocxParagraph(
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
                new DocxHyperlinkSpan("rIdStaticTextBox", null, null, null, "https://example.invalid/static-textbox", "External", null, 1, 1, 1, 1, 4)
            ]
        };
        DocxFloatingDrawing headerDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [headerDrawing]
            }
        };
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d);
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
            [new DocxParagraphElement(body)],
            [body],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/static-textbox", annotation.Uri);
        TestAssert.True(annotation.X > 70d, "Static floating text-box annotation should be anchored to the placed header drawing.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover static floating text-box hyperlink text.");
    }

    public static void DocxRendererEmitsFloatingTextBoxInternalHyperlinkDestinations()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Drawing target", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "TextBoxBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 14)
            ]
        };
        DocxParagraph targetParagraph = DocxTests.CreateDocxLayoutParagraph("Bookmark target", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("8", "TextBoxBookmarkTarget", 0, 0, 0)
            ]
        };
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(targetParagraph)]);
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
            [new DocxParagraphElement(linkParagraph)],
            [],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal floating text-box links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal links should resolve to floating text-box bookmark destinations.");
        TestAssert.True(annotation.Destination?.Left > 70d, "The destination should use placed floating text-box x coordinates.");
        TestAssert.True(annotation.Destination?.Top > 70d, "The destination should use placed floating text-box y coordinates.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should cover the rendered body internal-link text.");
    }

    public static void DocxRendererEmitsStaticFloatingTextBoxInternalHyperlinkDestinations()
    {
        DocxParagraph linkParagraph = new(
            [
                new DocxTextRun("Go ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Header target", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "StaticTextBoxBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 13)
            ]
        };
        DocxParagraph targetParagraph = DocxTests.CreateDocxLayoutParagraph("Bookmark target", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("13", "StaticTextBoxBookmarkTarget", 0, 0, 0)
            ]
        };
        DocxFloatingDrawing headerDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(targetParagraph)]);
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [headerDrawing]
            }
        };
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
            [new DocxParagraphElement(linkParagraph)],
            [linkParagraph],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal static floating text-box links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal links should resolve to static floating text-box bookmark destinations.");
        TestAssert.True(annotation.Destination?.Left > 70d, "The destination should use placed static floating text-box x coordinates.");
        TestAssert.True(annotation.Destination?.Top > 70d, "The destination should use placed static floating text-box y coordinates.");
        TestAssert.True(annotation.Width > 0d, "The clickable rectangle should cover the rendered body internal-link text.");
    }

    public static void DocxStaticStoryRendererEmitsInternalHyperlinkDestinations()
    {
        DocxParagraph header = new(
            [
                new DocxTextRun("Header ", 10d, null, false, false, false, null, null),
                new DocxTextRun("Jump", 10d, null, false, false, false, null, null)
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
            null)
        {
            Hyperlinks =
            [
                new DocxHyperlinkSpan(null, "StaticBookmarkTarget", null, null, null, null, null, 1, 1, 1, 1, 4)
            ]
        };
        DocxParagraph target = DocxTests.CreateDocxLayoutParagraph("Body target", 10d, 12d) with
        {
            BookmarkAnchors =
            [
                new DocxBookmarkAnchor("12", "StaticBookmarkTarget", 0, 0, 0)
            ]
        };
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [header]
            }
        };
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
            [new DocxParagraphElement(target)],
            [target],
            []);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.True(annotation.Uri is null, "Internal static-story links should not be emitted as URI actions.");
        TestAssert.True(annotation.Destination is { PageIndex: 0 }, "Internal static-story links should resolve through the shared bookmark destination map.");
        TestAssert.True(annotation.Y > 150d, "The clickable rectangle should be anchored to the header story text.");
        TestAssert.True(annotation.Destination?.Left >= document.MarginLeftPoints, "The destination should use placed body bookmark coordinates.");
        TestAssert.True(annotation.Width > 0d, "The annotation should cover static-story hyperlink text.");
    }
}
