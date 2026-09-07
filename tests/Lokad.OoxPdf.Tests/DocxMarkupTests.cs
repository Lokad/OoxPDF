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

internal static class DocxMarkupTests
{
    public static void DocxMarkupGeometryKeepsAuthoredMediaBox()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var finalRenderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var allRenderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        PdfPage finalPage = finalRenderer.RenderBlankPages(finalDocument, null, CancellationToken.None).Single();
        PdfPage allPage = allRenderer.RenderBlankPages(allDocument, null, CancellationToken.None).Single();
        DocxLayoutPageSnapshot finalLayoutPage = finalRenderer.InspectLayout(finalDocument).Pages.Single();
        DocxLayoutPageSnapshot allLayoutPage = allRenderer.InspectLayout(allDocument).Pages.Single();

        TestAssert.Equal(finalDocument.PageWidthPoints, finalPage.Width);
        TestAssert.Equal(finalDocument.PageHeightPoints, finalPage.Height);
        TestAssert.Equal(finalPage.Width, allPage.Width);
        TestAssert.Equal(finalPage.Height, allPage.Height);
        TestAssert.Equal(finalLayoutPage.Width, allLayoutPage.Width);
        TestAssert.Equal(finalLayoutPage.Height, allLayoutPage.Height);
        TestAssert.Equal(finalLayoutPage.ColumnFrameWidthSum, allLayoutPage.ColumnFrameWidthSum);
        TestAssert.Equal(finalLayoutPage.ColumnGutterWidthSum, allLayoutPage.ColumnGutterWidthSum);
        TestAssert.Equal(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout.ToString(), allRenderer.InspectLayout(allDocument).MarkupGeometryMode);
        TestAssert.Equal(0d, allLayoutPage.MarkupMarginReservePoints);
    }

    public static void DocxMarkupReserveMarginGeometryShrinksBodyAndPreservesMediaBox()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var preserveRenderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var reserveRenderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin);

        PdfPage preservePage = preserveRenderer.RenderBlankPages(allDocument, null, CancellationToken.None).Single();
        PdfPage reservePage = reserveRenderer.RenderBlankPages(allDocument, null, CancellationToken.None).Single();
        DocxLayoutSnapshot preserveLayout = preserveRenderer.InspectLayout(allDocument);
        DocxLayoutSnapshot reserveLayout = reserveRenderer.InspectLayout(allDocument);
        DocxLayoutPageSnapshot preserveLayoutPage = preserveLayout.Pages.Single();
        DocxLayoutPageSnapshot reserveLayoutPage = reserveLayout.Pages.Single();

        TestAssert.Equal(preservePage.Width, reservePage.Width);
        TestAssert.Equal(preservePage.Height, reservePage.Height);
        TestAssert.Equal(preserveLayoutPage.Width, reserveLayoutPage.Width);
        TestAssert.Equal(preserveLayoutPage.Height, reserveLayoutPage.Height);
        TestAssert.Equal(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin.ToString(), reserveLayout.MarkupGeometryMode);
        TestAssert.True(reserveLayout.MarkupMarginReservePoints > 0d, "Reserve-margin geometry should report the amount reserved for review balloons.");
        TestAssert.True(reserveLayoutPage.MarginRight > preserveLayoutPage.MarginRight, "Reserve-margin geometry should increase the effective right margin.");
        TestAssert.True(Math.Abs(reserveLayoutPage.MarginRight - 207d) < 0.001d, "Reserve-margin geometry should use the current Word-like Letter review margin target.");
        TestAssert.True(reserveLayoutPage.ColumnFrameWidthSum < preserveLayoutPage.ColumnFrameWidthSum, "Reserve-margin geometry should shrink the body column frame.");
        TestAssert.True(reserveRenderer.InspectMarkupBalloons(allDocument).All(placement => placement.X >= reserveLayoutPage.Width - reserveLayoutPage.MarginRight), "Right-side balloons should live in the reserved review margin.");
        TestAssert.True(reserveRenderer.InspectMarkupBalloons(allDocument).All(placement => placement.X > reserveLayoutPage.Width - reserveLayoutPage.MarginRight + 4d), "Right-side balloon bodies should leave a connector stem inside the reserved review margin.");
    }

    public static void DocxMarkupBalloonAreaClampsImpossibleReviewLaneInsideMediaBox()
    {
        const double nominalMinimumBalloonBodyWidth = 24d;
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Tight", fontSize: 6d, lineSpacingPoints: 7d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.Insertion, "tiny", "Reviewer", "2026-06-12T00:00:00Z", "ins", null, [])
            ]
        };
        DocxDocument document = new(
            20d,
            80d,
            6d,
            6d,
            6d,
            6d,
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
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxMarkupBalloonPlacementSnapshot placement = renderer
            .InspectMarkupBalloons(document)
            .Single(placement => !placement.IsOverflowSummary);

        TestAssert.True(
            placement.X >= 0d && placement.X + placement.Width <= document.PageWidthPoints,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Impossible review lanes should clamp balloon bodies inside the page media box. X={placement.X}, Width={placement.Width}, PageWidth={document.PageWidthPoints}."));
        TestAssert.True(
            placement.Width < nominalMinimumBalloonBodyWidth,
            "Pages narrower than the nominal review-lane minimum should shrink the fallback balloon body instead of emitting negative coordinates.");
        TestAssert.True(
            placement.BalloonConnectorX >= 0d && placement.BalloonConnectorX <= document.PageWidthPoints,
            "Impossible review lanes should keep the balloon connector inside the page media box.");
    }

    public static void DocxMarkupReserveMarginKeepsMirroredGutterInside()
    {
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First", fontSize: 10d, lineSpacingPoints: 10d);
        DocxParagraph second = first with
        {
            Runs = [new DocxTextRun("Second", 10d, null, false, false, false, null, null)]
        };
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            WidthValue = "12000",
            HeightValue = "6000",
            MarginLeftValue = "360",
            MarginRightValue = "360",
            MarginTopValue = "360",
            MarginBottomValue = "360",
            GutterDistanceValue = "240",
            GutterDistancePoints = 12d
        };
        var document = new DocxDocument(
            600d,
            300d,
            18d,
            18d,
            18d,
            18d,
            pageSettings,
            [],
            [],
            [],
            [
                new DocxParagraphElement(first),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(second)
            ],
            [first, second],
            [])
        {
            Settings = DocxDocumentSettings.Empty with
            {
                MirrorMargins = true
            }
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(30d, layout.Pages[0].MarginLeft);
        TestAssert.Equal(207d, layout.Pages[0].MarginRight);
        TestAssert.Equal(207d, layout.Pages[1].MarginLeft);
        TestAssert.Equal(30d, layout.Pages[1].MarginRight);
        TestAssert.Equal(189d, layout.Pages[0].MarkupMarginReservePoints);
        TestAssert.Equal(189d, layout.Pages[1].MarkupMarginReservePoints);
        TestAssert.Equal(30d, layout.Pages[0].ColumnFrames[0].X);
        TestAssert.Equal(207d, layout.Pages[1].ColumnFrames[0].X);
        TestAssert.Equal(layout.Pages[0].ColumnFrames[0].Width, layout.Pages[1].ColumnFrames[0].Width);
    }

    public static void DocxMarkupWordCompatibleGeometryUsesReserveMarginFallback()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var reserveRenderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin);
        var wordRenderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxLayoutSnapshot reserveLayout = reserveRenderer.InspectLayout(allDocument);
        DocxLayoutSnapshot wordLayout = wordRenderer.InspectLayout(allDocument);
        DocxTextEmissionSnapshot reserveEmission = reserveRenderer.InspectTextEmission(allDocument);
        DocxTextEmissionSnapshot wordEmission = wordRenderer.InspectTextEmission(allDocument);
        DocxLayoutPageSnapshot reserveLayoutPage = reserveLayout.Pages.Single();
        DocxLayoutPageSnapshot wordLayoutPage = wordLayout.Pages.Single();

        TestAssert.Equal(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup.ToString(), wordLayout.MarkupGeometryMode);
        TestAssert.Equal(reserveLayout.MarkupMarginReservePoints, wordLayout.MarkupMarginReservePoints);
        TestAssert.Equal(reserveLayoutPage.MarginRight, wordLayoutPage.MarginRight);
        TestAssert.Equal(reserveLayoutPage.ColumnFrameWidthSum, wordLayoutPage.ColumnFrameWidthSum);
        TestAssert.True(wordLayoutPage.TextLineHeightSum <= reserveLayoutPage.TextLineHeightSum + 0.001d, "Word-compatible all-markup should not increase aggregate text line metrics when the print-scale profile is active.");
        TestAssert.True(
            wordEmission.Lines.SelectMany(line => line.Segments).Max(segment => segment.PdfFontSize) <
            reserveEmission.Lines.SelectMany(line => line.Segments).Max(segment => segment.PdfFontSize),
            "Word-compatible all-markup should apply the current print-scale profile to emitted PDF font sizes.");
    }

    public static void DocxMarkupWordCompatibleGeometryHasNoEffectOutsideAllMarkup()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument simpleDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.SimpleMarkup);
        var finalRenderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.Final,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var simpleRenderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.SimpleMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxLayoutSnapshot finalLayout = finalRenderer.InspectLayout(finalDocument);
        DocxLayoutSnapshot simpleLayout = simpleRenderer.InspectLayout(simpleDocument);

        TestAssert.Equal(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout.ToString(), finalLayout.MarkupGeometryMode);
        TestAssert.Equal(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout.ToString(), simpleLayout.MarkupGeometryMode);
        TestAssert.Equal(0d, finalLayout.MarkupMarginReservePoints);
        TestAssert.Equal(0d, simpleLayout.MarkupMarginReservePoints);
    }

    public static void DocxWordCompatibleAllMarkupWrapsSpecialTokensInNarrowedBodyFrame()
    {
        DocxParagraph softHyphen = DocxTests.CreateDocxLayoutParagraph("ABCDEFGHIJKLMNOPQRST\u00ADUVWXYZABCDEFGHIJKLMNOPQRSTUV", 10d, 12d);
        DocxDocument softHyphenDocument = DocxTests.CreateAllMarkupWrapProbeDocument([softHyphen]);
        DocxTextLineLayout[] preserveSoftHyphen = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(softHyphenDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTextLineLayout>()
            .ToArray();
        DocxTextLineLayout[] wordSoftHyphen = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(softHyphenDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(1, preserveSoftHyphen.Length);
        TestAssert.Equal(2, wordSoftHyphen.Length);
        TestAssert.True(wordSoftHyphen[0].Text.EndsWith('\u00AD'), "Word-compatible all-markup should use the hidden soft hyphen when the review margin narrows an otherwise fitting token.");
        TestAssert.True(wordSoftHyphen[0].EndsWithIntraTokenBreak, "Soft-hyphen wraps in the narrowed review body should keep intra-token provenance.");

        DocxParagraph nonbreaking = DocxTests.CreateDocxLayoutParagraph("Alpha\u00A0Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa", 10d, 12d);
        DocxDocument nonbreakingDocument = DocxTests.CreateAllMarkupWrapProbeDocument([nonbreaking]);
        DocxTextLineLayout[] preserveNonbreaking = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(nonbreakingDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTextLineLayout>()
            .ToArray();
        DocxTextLineLayout[] wordNonbreaking = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(nonbreakingDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(1, preserveNonbreaking.Length);
        TestAssert.Equal(2, wordNonbreaking.Length);
        TestAssert.True(wordNonbreaking[0].Text.Contains("Alpha\u00A0Beta", StringComparison.Ordinal), "The narrowed review body should not split a nonbreaking-space token.");
        TestAssert.True(wordNonbreaking[1].Text.StartsWith("Theta", StringComparison.Ordinal), "The narrowed review body should wrap at the following breakable space.");

        DocxParagraph tabs = DocxTests.CreateDocxLayoutParagraph("Alpha\tBeta Gamma Delta Epsilon Zeta Eta Theta Iota", 10d, 12d);
        DocxDocument tabDocument = DocxTests.CreateAllMarkupWrapProbeDocument([tabs]);
        DocxTextLineLayout[] preserveTabs = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(tabDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTextLineLayout>()
            .ToArray();
        DocxTextLineLayout[] wordTabs = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(tabDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.Equal(1, preserveTabs.Length);
        TestAssert.Equal(2, wordTabs.Length);
        TestAssert.True(wordTabs[0].Text.Contains('\t'), "The first narrowed review line should retain the authored tab field.");
        TestAssert.True(wordTabs[1].Text.StartsWith("Theta", StringComparison.Ordinal), "The narrowed review body should wrap after the tab field without losing following text.");
        TestAssert.True(
            Math.Abs(wordTabs[0].Segments[1].X - (wordTabs[0].Segments[0].X + 36d)) < 0.001d,
            "Tab field text should keep the default-tab stop origin after all-markup narrowing.");

        DocxParagraph punctuation = DocxTests.CreateDocxLayoutParagraph("Alpha, beta; gamma: delta. Epsilon zeta eta theta iota kappa lambda.", 10d, 12d);
        DocxDocument punctuationDocument = DocxTests.CreateAllMarkupWrapProbeDocument([punctuation]);
        DocxTextLineLayout[] wordPunctuation = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(punctuationDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTextLineLayout>()
            .ToArray();

        TestAssert.True(wordPunctuation.Length >= 2, "The punctuation fixture should wrap in the narrowed all-markup review body.");
        TestAssert.True(
            wordPunctuation.Take(wordPunctuation.Length - 1).All(line => line.Width >= wordPunctuation.Last().Width),
            "Final narrowed all-markup lines should keep natural width while earlier punctuation-wrapped lines fill more of the body frame.");
        TestAssert.True(
            wordPunctuation.Any(line => line.Text.EndsWith('.') || line.Text.EndsWith(';') || line.Text.EndsWith(':') || line.Text.EndsWith(',')),
            "The narrowed review body should retain punctuation at line boundaries instead of dropping it during wrapping.");
    }

    public static void DocxWordCompatibleAllMarkupCapsLargeBodyTextAndTerminalAdvance()
    {
        DocxParagraph title = new(
            [new DocxTextRun("One-page markup margin geometry", 15d, null, true, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
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
            [new DocxParagraphElement(title)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        DocxTextEmissionSnapshot preserve = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectTextEmission(document);
        DocxTextEmissionSnapshot wordCompatible = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectTextEmission(document);

        DocxTextEmissionSegmentSnapshot[] preserveSegments = preserve.Lines.Single().Segments.ToArray();
        DocxTextEmissionSegmentSnapshot[] wordSegments = wordCompatible.Lines.Single().Segments.ToArray();
        double preserveVisibleFontSize = preserveSegments.Where(segment => !segment.IsTerminalLineSpace).Max(segment => segment.PdfFontSize);
        double wordVisibleFontSize = wordSegments.Where(segment => !segment.IsTerminalLineSpace).Max(segment => segment.PdfFontSize);
        DocxTextEmissionSegmentSnapshot preserveTerminal = preserveSegments.Single(segment => segment.IsTerminalLineSpace);
        DocxTextEmissionSegmentSnapshot wordTerminal = wordSegments.Single(segment => segment.IsTerminalLineSpace);

        TestAssert.True(
            wordVisibleFontSize < preserveVisibleFontSize - 1d,
            "Word-compatible all-markup should cap oversized body text to the Office-like heading font size.");
        TestAssert.True(
            wordTerminal.PdfFontSize < wordVisibleFontSize - 1d,
            "Word-compatible all-markup should emit terminal paragraph marks at body text size for capped heading lines.");
        TestAssert.True(
            wordTerminal.X < preserveTerminal.X - 25d,
            "Word-compatible all-markup should place terminal paragraph marks at the capped emitted text advance.");
    }

    public static void DocxWordCompatibleAllMarkupScalesBodySpacing()
    {
        DocxParagraph first = new(
            [new DocxTextRun("First paragraph", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Second paragraph", 11d, 13.2d);
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
            [new DocxParagraphElement(first), new DocxParagraphElement(second)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxLayoutItemSnapshot reserveSecond = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .InspectLayout(document)
            .Pages.Single()
            .Items.Where(item => item.Kind == "TextLine")
            .Skip(1)
            .First();
        DocxLayoutItemSnapshot wordSecond = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectLayout(document)
            .Pages.Single()
            .Items.Where(item => item.Kind == "TextLine")
            .Skip(1)
            .First();

        TestAssert.Equal(8d, reserveSecond.AppliedBeforeSpacingPoints ?? -1d);
        TestAssert.True(
            Math.Abs((wordSecond.AppliedBeforeSpacingPoints ?? 0d) - 6.739d) < 0.001d,
            "Word-compatible all-markup should scale paragraph spacing with the print profile.");
    }

    public static void DocxWordCompatibleAllMarkupScalesFloatingTextBoxSpacing()
    {
        DocxParagraph first = new(
            [new DocxTextRun("Text box first", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Text box second", 11d, 13.2d);
        DocxFloatingDrawing drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(first), new DocxParagraphElement(second)]);
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
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Anchor", 11d, 13.2d))],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxTextLineLayout reserveSecond = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .FloatingDrawings.Single()
            .TextBoxLayout!.TextLines.Skip(1).First();
        DocxTextLineLayout wordSecond = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .FloatingDrawings.Single()
            .TextBoxLayout!.TextLines.Skip(1).First();

        TestAssert.Equal(8d, reserveSecond.AppliedBeforeSpacing ?? -1d);
        TestAssert.True(
            Math.Abs((wordSecond.AppliedBeforeSpacing ?? 0d) - 6.739d) < 0.001d,
            "Word-compatible all-markup should scale floating text-box paragraph spacing with the print profile.");
    }

    public static void DocxWordCompatibleAllMarkupScalesStaticTextBoxSpacing()
    {
        DocxParagraph first = new(
            [new DocxTextRun("Static text box first", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Static text box second", 11d, 13.2d);
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(first), new DocxParagraphElement(second)])]
            }
        };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            pageSettings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 11d, 13.2d))],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxTextLineLayout reserveSecond = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .StaticFloatingDrawings.Single()
            .TextBoxLayout!.TextLines.Skip(1).First();
        DocxTextLineLayout wordSecond = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .StaticFloatingDrawings.Single()
            .TextBoxLayout!.TextLines.Skip(1).First();

        TestAssert.Equal(8d, reserveSecond.AppliedBeforeSpacing ?? -1d);
        TestAssert.True(
            Math.Abs((wordSecond.AppliedBeforeSpacing ?? 0d) - 6.739d) < 0.001d,
            "Word-compatible all-markup should scale static text-box paragraph spacing with the print profile.");
    }

    public static void DocxWordCompatibleAllMarkupAddsBodyPositioningSpacing()
    {
        DocxParagraph paragraph = new(
            [
                new DocxTextRun("Body text ", 11d, null, false, false, false, null, null),
                new DocxTextRun("operation spacing", 11d, null, false, false, false, null, null)
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
            null);
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

        DocxTextEmissionSegmentSnapshot[] preserveVisible = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectTextEmission(document)
            .Lines.Single()
            .Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();
        DocxTextEmissionSegmentSnapshot[] wordVisible = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectTextEmission(document)
            .Lines.Single()
            .Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();
        DocxTextEmissionSegmentSnapshot preserveSegment = preserveVisible[0];
        DocxTextEmissionSegmentSnapshot wordSegment = wordVisible[0];

        TestAssert.True(Math.Abs(preserveSegment.PositioningCharacterSpacing) < 0.0001d, "Preserve layout should not invent positioned body tracking.");
        TestAssert.Equal("None", wordSegment.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(wordSegment.PdfCharacterSpacing) < 0.0001d, "Word-compatible body tracking should stay out of PDF Tc.");
        TestAssert.True(
            Math.Abs(wordSegment.PositioningCharacterSpacing - 0.071d) < 0.0001d,
            "Word-compatible all-markup should emit the Office-like body tracking through positioned glyph advances.");
        TestAssert.True(
            wordSegment.AdvanceProfile.PositioningCharacterSpacingGapTotal > 0.5d,
            "Text-emission snapshots should expose the positioned body tracking contribution.");
        TestAssert.True(
            Math.Abs(wordSegment.AdvanceProfile.TextStateCharacterSpacingGapTotal) < 0.0001d,
            "Word-compatible body tracking should not become PDF text-state spacing.");
        TestAssert.True(
            wordVisible[1].X - wordVisible[0].X < preserveVisible[1].X - preserveVisible[0].X - 0.2d,
            "Word-compatible all-markup should shift later body text operations toward Office-like emitted x origins.");
    }

    public static void DocxWordCompatibleAllMarkupUsesMoveRevisionPositioningProfile()
    {
        var deletionRevision = new DocxRevisionInfo(DocxRevisionKind.Deletion, "3", "Reviewer", "2026-06-10T00:00:00Z", "del", null, []);
        var moveFromRevision = new DocxRevisionInfo(DocxRevisionKind.MoveFrom, "1", "Reviewer", "2026-06-10T00:00:00Z", "moveFrom", null, []);
        var insertionRevision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "4", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        var moveToRevision = new DocxRevisionInfo(DocxRevisionKind.MoveTo, "2", "Reviewer", "2026-06-10T00:00:00Z", "moveTo", null, []);
        DocxParagraph deletedParagraph = CreateRevisedParagraph("ChangedFrom", deletionRevision);
        DocxParagraph movedFromParagraph = CreateRevisedParagraph("ChangedFrom", moveFromRevision);
        DocxParagraph insertedParagraph = CreateRevisedParagraph("ChangedTo", insertionRevision);
        DocxParagraph movedToParagraph = CreateRevisedParagraph("ChangedTo", moveToRevision);
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
                new DocxParagraphElement(deletedParagraph),
                new DocxParagraphElement(movedFromParagraph),
                new DocxParagraphElement(insertedParagraph),
                new DocxParagraphElement(movedToParagraph)
            ],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxTextEmissionLineSnapshot[] wordLines = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectTextEmission(document)
            .Lines
            .ToArray();
        DocxTextEmissionSegmentSnapshot wordDeletion = RevisionSegment(wordLines, "Deletion");
        DocxTextEmissionSegmentSnapshot wordMoveFrom = RevisionSegment(wordLines, "MoveFrom");
        DocxTextEmissionSegmentSnapshot wordInsertion = RevisionSegment(wordLines, "Insertion");
        DocxTextEmissionSegmentSnapshot wordMoveTo = RevisionSegment(wordLines, "MoveTo");

        TestAssert.True(Math.Abs(wordMoveFrom.PositioningCharacterSpacing - 0.060d) < 0.0001d, "Word-compatible moved-from text should use the deleted-text positioning profile.");
        TestAssert.True(Math.Abs(wordMoveTo.PositioningCharacterSpacing - 0.126d) < 0.0001d, "Word-compatible moved-to text should use the inserted-text positioning profile.");
        TestAssert.True(Math.Abs(wordMoveFrom.X - wordDeletion.X) < 0.001d, "Word-compatible moved-from text should share the deleted-text X positioning branch.");
        TestAssert.True(Math.Abs(wordMoveTo.X - wordInsertion.X) < 0.001d, "Word-compatible moved-to text should share the inserted-text X positioning branch.");

        static DocxParagraph CreateRevisedParagraph(string revisedText, DocxRevisionInfo revision)
        {
            return new DocxParagraph(
                [
                    new DocxTextRun("Prefix ", 10d, null, false, false, false, null, null),
                    new DocxTextRun(revisedText, 10d, null, false, false, false, null, null)
                    {
                        Revision = revision
                    }
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
                null);
        }

        static DocxTextEmissionSegmentSnapshot RevisionSegment(
            IEnumerable<DocxTextEmissionLineSnapshot> lines,
            string revisionKind)
        {
            return lines
                .SelectMany(line => line.Segments)
                .Single(segment => segment.RevisionKind == revisionKind && !segment.IsTerminalLineSpace);
        }
    }

    public static void DocxWordCompatibleAllMarkupUsesMoveRevisionPositioningProfileAcrossTextFlows()
    {
        DocxParagraph floatingTextBoxParagraph = CreateMoveRevisionParagraph();
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(floatingTextBoxParagraph)]);
        DocxDocument floatingDocument = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d))],
            []) with
        {
            FloatingDrawings = [floatingDrawing],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        AssertMoveRevisionStoryProfile(
            floatingDocument,
            line => line.StoryKind == "TextBox" && line.ContainerStoryKind == "Body",
            "body-anchored floating text box");

        DocxParagraph staticTextBoxParagraph = CreateMoveRevisionParagraph();
        DocxPageSettings staticSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(staticTextBoxParagraph)])]
            }
        };
        DocxDocument staticDocument = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            []) with
        {
            PageSettings = staticSettings,
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        AssertMoveRevisionStoryProfile(
            staticDocument,
            line => line.IsStaticStory && line.StoryKind == "TextBox" && line.ContainerStoryKind == "Header",
            "static header text box");

        DocxParagraph footnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "91",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "91",
            [new DocxParagraphElement(CreateMoveRevisionParagraph())],
            [],
            [], null);
        DocxDocument footnoteDocument = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(footnoteAnchor)], [])
            with
            {
                RelatedStories = [footnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
        AssertMoveRevisionStoryProfile(
            footnoteDocument,
            line => line.StoryKind == "Footnote",
            "placed footnote");

        DocxParagraph endnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "92",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "92",
            [new DocxParagraphElement(CreateMoveRevisionParagraph())],
            [],
            [], null);
        DocxDocument endnoteDocument = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(endnoteAnchor)], [])
            with
            {
                RelatedStories = [endnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };
        AssertMoveRevisionStoryProfile(
            endnoteDocument,
            line => line.StoryKind == "Endnote",
            "placed endnote");

        static DocxParagraph CreateMoveRevisionParagraph()
        {
            var moveFromRevision = new DocxRevisionInfo(DocxRevisionKind.MoveFrom, "11", "Reviewer", "2026-06-10T00:00:00Z", "moveFrom", null, []);
            var moveToRevision = new DocxRevisionInfo(DocxRevisionKind.MoveTo, "12", "Reviewer", "2026-06-10T00:00:00Z", "moveTo", null, []);
            return new DocxParagraph(
                [
                    new DocxTextRun("Before ", 10d, null, false, false, false, null, null),
                    new DocxTextRun("MovedFrom", 10d, null, false, false, false, null, null)
                    {
                        Revision = moveFromRevision
                    },
                    new DocxTextRun(" After ", 10d, null, false, false, false, null, null),
                    new DocxTextRun("ChangedTo", 10d, null, false, false, false, null, null)
                    {
                        Revision = moveToRevision
                    }
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
                null);
        }

        static void AssertMoveRevisionStoryProfile(
            DocxDocument document,
            Func<DocxTextEmissionLineSnapshot, bool> linePredicate,
            string flowName)
        {
            DocxTextEmissionSegmentSnapshot[] segments = new DocxRenderer(
                    fontResolver: null,
                    markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                    markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .InspectTextEmission(document)
                .Lines
                .Where(linePredicate)
                .SelectMany(line => line.Segments)
                .Where(segment => !segment.IsTerminalLineSpace)
                .ToArray();
            DocxTextEmissionSegmentSnapshot moveFrom = segments.Single(segment => segment.RevisionKind == "MoveFrom");
            DocxTextEmissionSegmentSnapshot moveTo = segments.Single(segment => segment.RevisionKind == "MoveTo");

            TestAssert.True(
                Math.Abs(moveFrom.PositioningCharacterSpacing - 0.060d) < 0.0001d,
                "Word-compatible moved-from text should use the deleted-text positioning profile in " + flowName + ".");
            TestAssert.True(
                Math.Abs(moveTo.PositioningCharacterSpacing - 0.126d) < 0.0001d,
                "Word-compatible moved-to text should use the inserted-text positioning profile in " + flowName + ".");
        }
    }

    public static void DocxWordCompatibleAllMarkupPaintsReviewLaneBackground()
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

        TestAssert.Contains("0.949 g", page.Content);
        TestAssert.Contains("411.93 89.475 199.7 614.25 re f", page.Content);
    }

    public static void DocxWordCompatibleAllMarkupMirrorsReviewLaneBackground()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "docx-markup-margin-mirrored.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        PdfPage[] pages = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .ToArray();

        TestAssert.True(pages.Length >= 2, "The mirrored-margin fixture should render odd and even review pages.");
        TestAssert.Contains("411.93 89.475 199.7 614.25 re f", pages[0].Content);
        TestAssert.Contains("0.37 89.475 199.7 614.25 re f", pages[1].Content);
    }

    public static void DocxWordCompatibleAllMarkupUsesOfficeRevisionDecorationColor()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0.82 0.204 0.22 rg", page.Content);
        TestAssert.Contains("0.475 re f", page.Content);
    }

    public static void DocxWordCompatibleAllMarkupEmitsOfficeLikeMarkupPrimitiveInventory()
    {
        var insertion = new DocxRevisionInfo(DocxRevisionKind.Insertion, "1", "Reviewer", "2026-06-01T00:00:00Z", "ins", null, []);
        DocxTextRun commentRun = new("Commented", 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        DocxParagraph markupParagraph = new(
            [
                commentRun,
                new DocxTextRun(" inserted", 10d, null, false, false, true, "single", null) { Revision = insertion },
                new DocxTextRun(" double", 10d, null, false, false, true, "double", null),
                new DocxTextRun(" strike", 10d, null, false, false, false, null, null, 0d, false, null, true, null, false, null, null, null, null, null, false, null, false, null, null),
                new DocxTextRun(" double-strike", 10d, null, false, false, false, null, null, 0d, false, null, false, null, true, null, null, null, null, null, false, null, false, null, null)
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
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 0, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ],
            CommentRanges =
            [
                new DocxCommentRange("1", 0, 0, 1, commentRun.Text.Length, 1, 0)
            ],
            Revisions = [insertion]
        };
        DocxParagraph tableParagraph = DocxTests.CreateDocxLayoutParagraph("Bordered", 10d, 12d);
        var cell = new DocxTableCell(
            "Bordered",
            [tableParagraph],
            FillHex: null,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders:
            [
                new DocxTableCellBorder("top", "single", "0000FF", "8"),
                new DocxTableCellBorder("bottom", "double", "0000FF", "8")
            ],
            DocxTableCellMargins.Empty);
        var table = new DocxTable(LayoutValue: null, ColumnWidthsPoints: [120d], Rows: [new DocxTableRow([cell], HeightPoints: 24d)]);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public comment body", 10d, 12d))],
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
            [new DocxParagraphElement(markupParagraph), new DocxTableElement(table)],
            [],
            [table])
        {
            RelatedStories = [commentStory],
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
        TestAssert.Contains("0.475 w", page.Content);
        TestAssert.Contains("[0.475 0.475] 0 d", page.Content);
        TestAssert.Contains("27.925 ", page.Content);
        TestAssert.Contains("0 0 1 rg", page.Content);
        TestAssert.True(DocxTests.CountOccurrences(page.Content, " re B*") >= 1, "Markup balloons should render as fill/stroke rectangle primitives.");
        TestAssert.True(DocxTests.CountOccurrences(page.Content, " l S") >= 6, "Comment ranges and balloon connectors should render stroked line primitives.");
        TestAssert.True(DocxTests.CountOccurrences(page.Content, " re f") >= 8, "Decorations, change bars, comment markers, and table borders should render filled rectangle primitives.");
    }

    public static void DocxWordCompatibleAllMarkupRendersRevisionBalloonTextWithOfficeFont()
    {
        // Office A/B: Word balloons formatting revisions (insertions/deletions/moves render
        // inline), so the font probe below uses a formatting revision.
        DocxParagraph formatParagraph = DocxTests.CreateDocxLayoutParagraph("Formatted revision anchor", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(DocxRevisionKind.RunPropertiesChange, "1", "Reviewer", "2026-06-01T00:00:00Z", "rPrChange", null, propertyElementNames: ["b"])
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
            [new DocxParagraphElement(formatParagraph)],
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

        TestAssert.True(
            DocxTests.CountOccurrences(page.Content, "6.975 Tf") >= 2,
            "Word-compatible revision balloons should render Office-sized title and body text.");
        TestAssert.True(
            !page.Content.Contains(" 5.5 Tf", StringComparison.Ordinal),
            "Word-compatible revision balloons should not use the legacy summary title font size.");
        TestAssert.True(
            !page.Content.Contains(" 5 Tf", StringComparison.Ordinal),
            "Word-compatible revision balloons should not use the legacy summary body font size.");
    }

    public static void DocxWordCompatibleAllMarkupWrapsLongRevisionBalloonBody()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Long revision summary anchor", 10d, 12d) with
        {
            Revisions =
            [
                new DocxRevisionInfo(
                    DocxRevisionKind.ParagraphPropertiesChange,
                    "1",
                    "Reviewer",
                    "2026-06-01T00:00:00Z",
                    "pPr",
                    null,
                    propertyElementNames:
                    [
                        "keepNext",
                        "keepLines",
                        "pageBreakBefore",
                        "widowControl",
                        "contextualSpacing",
                        "spacing",
                        "ind",
                        "jc",
                        "tabs"
                    ]),
                new DocxRevisionInfo(DocxRevisionKind.Deletion, "2", "Reviewer", "2026-06-01T00:00:00Z", "del", null, []),
                new DocxRevisionInfo(DocxRevisionKind.MoveFrom, "3", "Reviewer", "2026-06-01T00:00:00Z", "moveFrom", null, [])
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

        DocxMarkupBalloonPlacementSnapshot placement = renderer.InspectMarkupBalloons(document)
            .Single(item => item.Kind == "Revision");
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(1, placement.WordCompatibleBodySummaryPartCount);
        TestAssert.True(
            DocxTests.CountOccurrences(page.Content, "6.975 Tf") >= 4,
            "Long Word-compatible revision balloons should render title, first body line, continuation line, and terminal spacing at the Office-sized font.");
        TestAssert.True(
            DocxTests.CountOccurrences(page.Content, "] TJ") >= 2,
            "Long Word-compatible revision balloons should use positioned glyph arrays for the title and wrapped continuation body line.");
        TestAssert.True(
            !page.Content.Contains(" 5.5 Tf", StringComparison.Ordinal) &&
            !page.Content.Contains(" 5 Tf", StringComparison.Ordinal),
            "Wrapped Word-compatible revision balloons should stay on the Office text profile instead of the legacy fallback.");
    }

    public static void DocxSimpleMarkupRendererDrawsChangeBars()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.SimpleMarkup);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot(document.Paragraphs.Single().Revisions), "rg"), page.Content);
        TestAssert.Contains(" 1.5 ", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsChangeBarsForRunRevisions()
    {
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "13", "Reviewer", "2026-06-01T00:00:00Z", "ins", null, []);
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Simple run revision", 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun("Simple ", 10d, null, false, false, false, null, null),
                new DocxTextRun("run", 10d, null, false, false, false, null, null)
                {
                    Revision = revision
                },
                new DocxTextRun(" revision", 10d, null, false, false, false, null, null)
            ]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []) with
        {
            MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
        };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot([revision]), "rg"), page.Content);
        TestAssert.Contains(" 1.5 ", page.Content);
    }

    public static void DocxSimpleMarkupSuppressesDeletedTextAndBalloons()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.SimpleMarkup);
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        DocxTextEmissionSnapshot textEmission = renderer.InspectTextEmission(document);
        DocxMarkupBalloonPlacementSnapshot[] balloons = renderer.InspectMarkupBalloons(document).ToArray();

        TestAssert.Equal("Before Inserted MovedTo After", string.Concat(document.Paragraphs.Single().Runs.Select(run => run.Text)));
        TestAssert.True(layout.RevisionItemCount >= 1, "Simple markup should retain changed-line candidates for margin change bars.");
        TestAssert.True(textEmission.InsertionRevisionSegmentCount >= 1, "Simple markup should keep inserted final-view text emission provenance.");
        TestAssert.True(textEmission.MoveToRevisionSegmentCount >= 1, "Simple markup should keep moved-to final-view text emission provenance.");
        TestAssert.Equal(0, textEmission.DeletionRevisionSegmentCount);
        TestAssert.Equal(0, textEmission.MoveFromRevisionSegmentCount);
        TestAssert.Equal(0, balloons.Length);
    }

    public static void DocxSimpleMarkupRendererDrawsPlacedFootnoteChangeBars()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "19",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "19", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote inserted text", 10d, 12d) with
        {
            Revisions = [revision]
        };
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "19",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [footnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
            };
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(
            layout.Pages.SelectMany(page => page.PlacedRelatedItems).Any(item => item.RevisionCount != 0),
            "Simple-markup layout should retain placed footnote revision candidates for margin change bars.");
        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot([revision]), "rg"), page.Content);
        TestAssert.Contains(" 1.5 ", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsPlacedEndnoteChangeBars()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "29",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "29", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote inserted text", 10d, 12d) with
        {
            Revisions = [revision]
        };
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "29",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [endnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
            };
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(
            layout.Pages.SelectMany(page => page.PlacedRelatedItems).Any(item => item.RevisionCount != 0),
            "Simple-markup layout should retain placed endnote revision candidates for margin change bars.");
        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot([revision]), "rg"), page.Content);
        TestAssert.Contains(" 1.5 ", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsFloatingTextBoxChangeBars()
    {
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "31", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph("Floating text box revision", 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun("Floating ", 10d, null, false, false, false, null, null),
                new DocxTextRun("revision", 10d, null, false, false, false, null, null)
                {
                    Revision = revision
                }
            ]
        };
        DocxFloatingDrawing drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d))],
            []) with
        {
            FloatingDrawings = [drawing],
            MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
        };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot([revision]), "rg"), page.Content);
        TestAssert.Contains(" 1.5 ", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsStaticFloatingTextBoxChangeBars()
    {
        var revision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "32", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph("Static text box revision", 10d, 12d) with
        {
            Runs =
            [
                new DocxTextRun("Static ", 10d, null, false, false, false, null, null),
                new DocxTextRun("revision", 10d, null, false, false, false, null, null)
                {
                    Revision = revision
                }
            ]
        };
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)])]
            }
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([], []) with
        {
            PageSettings = settings,
            MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
        };

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot([revision]), "rg"), page.Content);
        TestAssert.Contains(" 1.5 ", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsCommentMarkers()
    {
        string input = DocxTests.WriteCommentMarkerProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.SimpleMarkup);

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Contains("1 0.753 0 rg", page.Content);
        TestAssert.Contains("0.851 0.592 0 RG", page.Content);
        TestAssert.Contains(" re f", page.Content);
        TestAssert.Contains(" re S", page.Content);
        TestAssert.Contains(" Tj", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsCommentMarkersAcrossTextFlows()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateCommentMarkerParagraph("Body marker", "1");
        DocxTests.AssertSimpleMarkupCommentMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(bodyParagraph)], [], DocxPageSettings.Empty, []),
            "body text");

        DocxParagraph tableParagraph = DocxTests.CreateCommentMarkerParagraph("Table marker", "2");
        var tableCell = new DocxTableCell(
            "Table marker",
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
        DocxTests.AssertSimpleMarkupCommentMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxTableElement(table)], [], DocxPageSettings.Empty, []),
            "table cell text");

        DocxParagraph headerParagraph = DocxTests.CreateCommentMarkerParagraph("Header marker", "3");
        DocxPageSettings headerSettings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(headerParagraph)]
            }
        };
        DocxTests.AssertSimpleMarkupCommentMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([], [], headerSettings, []),
            "static header text");

        DocxParagraph drawingAnchor = DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateCommentMarkerParagraph("Text box marker", "4");
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
        DocxTests.AssertSimpleMarkupCommentMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(drawingAnchor)], [floatingDrawing], DocxPageSettings.Empty, []),
            "floating text-box text");

        DocxParagraph staticTextBoxParagraph = DocxTests.CreateCommentMarkerParagraph("Static text box marker", "5");
        DocxPageSettings staticFloatingSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(staticTextBoxParagraph)])]
            }
        };
        DocxTests.AssertSimpleMarkupCommentMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([], [], staticFloatingSettings, []),
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
        DocxParagraph footnoteParagraph = DocxTests.CreateCommentMarkerParagraph("Footnote marker", "6");
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "21",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        DocxTests.AssertSimpleMarkupCommentMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(footnoteAnchor)], [], DocxPageSettings.Empty, []) with
            {
                RelatedStories = [footnoteStory]
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
        DocxParagraph endnoteParagraph = DocxTests.CreateCommentMarkerParagraph("Endnote marker", "7");
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "22",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxTests.AssertSimpleMarkupCommentMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(endnoteAnchor)], [], DocxPageSettings.Empty, []) with
            {
                RelatedStories = [endnoteStory]
            },
            "placed endnote text");
    }
}
