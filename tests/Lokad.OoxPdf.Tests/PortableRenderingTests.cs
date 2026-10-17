using System.IO.Compression;
using System.Xml.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

// Real rendering with one pinned, redistributable face. These are content and
// resource checks on both platforms; Office visual comparisons remain separate.
internal static class PortableRenderingTests
{
    public static void PortableRenderingMathRemaindersKeepMixedStoryContentAndInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "remainder-placement-office.json", splitMathRuns: true);

    public static void PortableRenderingEmptyMathBodiesKeepFunctionAndOperatorInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "math-body-office.json", splitMathRuns: true);

    public static void PortableRenderingStructuralPaintRetainsVisibleSymbolsAndInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "math-source-office.json", splitMathRuns: true);

    public static void PortableRenderingStructuralPaintRetainsWorkBounds() =>
        DocxInlineCompositionTests.CheckStructuralPaintBounds(new PinnedResolver(mathFont: true));

    public static void PortableRenderingStructuralPaintKeepsMixedStoryContentAndInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "structural-placement-office.json", splitMathRuns: true);

    public static void PortableRenderingStructuralOnlyRevisionsRetainNeutrality() =>
        DocxInlineCompositionTests.CheckStructuralOnlyRevisionNeutrality(new PinnedResolver(mathFont: true));

    public static void PortableRenderingWholeStructuralRevisionsKeepContentAndInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "whole-math-office.json", splitMathRuns: true);

    public static void PortableRenderingEmptyRootDegreesKeepVisibleRootBodyAndInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "math-degree-office.json", splitMathRuns: true);

    public static void PortableRenderingMathRemaindersKeepVisibleArgumentsAndInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "math-remaining-office.json", splitMathRuns: true);

    public static void PortableRenderingEmptyRolesKeepMixedStoryContentAndInk() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "empty-role-placement-office.json", splitMathRuns: true);

    public static void PortableRenderingMathRolesKeepRemainingArgumentsAndInk() =>
        DocxInlineCompositionTests.CheckMathRoles(new PinnedResolver(mathFont: true), false);

    public static void PortableRenderingMathRunRevisionsKeepIndependentGlyphOwners() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false,
            "math-run-revisions-office.json", splitMathRuns: true);

    public static void PortableRenderingEmptyFractionsKeepStructureWithoutGlyphs() =>
        DocxInlineCompositionTests.CheckEmptyFractions(new PinnedResolver(mathFont: true), false);

    public static void PortableRenderingEquationRevisionsKeepVisibleContentAndOwners() =>
        DocxInlineCompositionTests.CheckEquationRevisions(new PinnedResolver(mathFont: true), false);

    public static void PortableRenderingRevisionNeutralityKeepsGlyphsAndPlacement() =>
        DocxRevisionNeutralityTests.CheckRendering(new PinnedResolver(), false);

    public static void PortableRenderingEmptyRunOwnersKeepMathAndAuthorOrder() =>
        DocxCommentsTests.CheckEmptyRevisionOwners(new PinnedResolver(mathFont: true));

    public static void PortableRenderingPictureRevisionBodyCellFramesKeepSourceDimensions() =>
        DocxPictureRevisionTests.CheckBodyCellPictureFrames(new PinnedResolver());

    public static void PortableRenderingPictureRevisionStaticFramesKeepVisibleContent() =>
        DocxPictureRevisionTests.CheckStaticPictureFrames(new PinnedResolver(), false);

    public static void PortableRenderingPictureRevisionSharedRunsKeepBothImagesAndText() =>
        DocxPictureRevisionTests.CheckSharedPictures(new PinnedResolver());

    public static void PortableRenderingPictureRevisionMarkersKeepSourceOwners() =>
        DocxPictureRevisionTests.CheckRendering(new PinnedResolver(), false);

    public static void PortableRenderingRevisionAuthorInkMatchesWordSourceOrder() =>
        DocxCommentsTests.CheckReviewRevisionInk(new PinnedResolver());

    public static void PortableRenderingNonTextReviewAuthorsKeepSourceOwners() =>
        DocxCommentsTests.CheckNonTextReviewAuthorColors(new PinnedResolver());

    public static void PortableRenderingMarkFormattingBalloonsKeepSourceOwners() =>
        DocxInlineCompositionTests.CheckMarkFormattingBalloons(new PinnedResolver(mathFont: true));

    public static void PortableRenderingNumberingContinuationKeepsVisibleContent() =>
        DocxInlineCompositionTests.CheckNumberingCounterContent(new PinnedResolver(mathFont: true), "numbering-continuations-office.json");

    public static void PortableRenderingNumberingCounterPaintsKeepVisibleContent() =>
        DocxInlineCompositionTests.CheckNumberingCounterContent(new PinnedResolver(mathFont: true));

    public static void PortableRenderingNumberingMarkDecorationsKeepSourceOwners() =>
        DocxInlineCompositionTests.CheckNumberingMarkDecorations(new PinnedResolver(mathFont: true), render: true);

    public static void PortableRenderingNumberingMarkOwnersKeepVisibleContent() =>
        DocxInlineCompositionTests.CheckNumberingMarkContent(new PinnedResolver(mathFont: true));

    public static void PortableRenderingUnchangedListDecorationsKeepSourceOwners() =>
        DocxInlineCompositionTests.CheckUnchangedListDecorations(new PinnedResolver());

    public static void PortableRenderingDefaultMathFontCoversItalicAlphabet()
    {
        var resolver = new PinnedResolver(mathFont: true);
        var resolution = resolver.Resolve(new FontRequest("Cambria Math"));
        var font = OpenTypeFont.Load(resolution.Source.GetBytesAsync().GetAwaiter().GetResult().ToArray());
        TestAssert.True(font.HasTrueTypeOutlines && font.TryGetMathConstants(out _),
            "The pinned math face must supply managed TrueType outlines and MATH constants.");
        foreach (int scalar in new[] { 0x1D44E, 0x1D44F, 0x1D450, 0x1D465, 0x1D466 })
            TestAssert.True(font.MapCodePoint(scalar) != 0,
                "Constants alone do not establish mathematical italic glyph coverage: " + scalar.ToString("X"));
    }

    public static void PortableRenderingDefaultMathKeepsTextRevisionOwners() =>
        DocxInlineCompositionTests.CheckTextRevisionMath(new PinnedResolver(mathFont: true), preserveDefaultMath: true);

    public static void PortableRenderingDefaultMathComposesAcrossStories()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-run-format-" + story + "-bold.docx"), mode);
            var paragraphs = DocxBlockTraversal.EnumerateSourceParagraphs(document).ToArray();
            var owners = paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(new PinnedResolver(mathFont: true), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).Single();
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(font => font.Font).ToHashSet();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                .Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(18, shows.Length);
            TestAssert.True(shows.All(show => show.X >= 0d && show.X < page.Width && show.Y > 0d && show.Y < page.Height),
                "Ordinary mathematical italic content must reach its visible body/cell/static frame.");
            string math = string.Concat(shows.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, math.Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(2, math.Split("a+b=c", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(1, page.Images.Count);
            TestAssert.Contains("share the same input series.", DocxPageFootnoteNumberingTests.DecodeText(page));
            string part = story == "body" || story == "cell" ? document.MainPartName : "/word/" + story + "1.xml";
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION" || d.PartName != part),
                "A supplied complete math face cannot silently lose supported selected-story equations.");
            TestAssert.True(owners.Zip(paragraphs.SelectMany(paragraph => paragraph.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Portable default math retains authored runs and source ownership.");
            CheckWrittenPages([page]);
        }
    }

    public static void PortableRenderingUndeclaredTypefacesKeepSourceOwners() =>
        DocxInlineCompositionTests.CheckDefaultTypefaces(new PinnedResolver());

    public static void PortableRenderingListLabelFamiliesKeepSourceOwners() =>
        DocxInlineCompositionTests.CheckListLabelFamilies(new PinnedResolver());

    public static void PortableRenderingListLabelStylesKeepSourceOwners() =>
        DocxInlineCompositionTests.CheckListLabelStyles(new PinnedResolver());

    public static void PortableRenderingStaticMissingMathKeepsSourceContent() =>
        DocxInlineCompositionTests.CheckStaticMissingMath(new PinnedResolver());

    public static void PortableRenderingStaticListsKeepLabelsAcrossModes() =>
        DocxInlineCompositionTests.CheckStaticListLabels(new PinnedResolver());

    public static void PortableRenderingStaticSourceFramesKeepProseOnPage() =>
        DocxInlineCompositionTests.CheckStaticSourceFrames(new PinnedResolver());

    public static void PortableRenderingMeasuredFootnoteOwnersKeepLabelsAndDestinations() =>
        DocxPageFootnoteNumberingTests.CheckMeasuredFootnoteOwners(new PinnedResolver());

    public static void PortableRenderingFormattingBalloonDescriptionsKeepSourceOwners() =>
        DocxInlineCompositionTests.CheckRunFormattingBalloonDescriptions(new PinnedResolver());

    public static void PortableRenderingRunFormattingRetainsAuthoredInk() =>
        DocxInlineCompositionTests.CheckRunFormattingInk(new PinnedResolver());

    public static void PortableRenderingRunFormattingMathKeepsSourceAndMetadata() =>
        DocxInlineCompositionTests.CheckRunFormattingMath(new PinnedResolver());

    public static void PortableRenderingReviewAuthorsShareSourceTextPalette() =>
        DocxCommentsTests.CheckReviewAuthorColors(new PinnedResolver());

    public static void PortableRenderingRevisionBarsKeepIndependentPageFrames() =>
        DocxRevisionBarTests.CheckHorizontalFrames(new PinnedResolver());

    public static void PortableRenderingTextRevisionMathKeepsSourceOwnersAndFallbacks() =>
        DocxInlineCompositionTests.CheckTextRevisionMath(new PinnedResolver());

    public static void PortableRenderingNumberedExplicitMathRetainsAuthoredRowOwnership() =>
        DocxInlineCompositionTests.CheckNumberedDisplayRows(new PinnedResolver());

    public static void PortableRenderingPictureFirstTablesRetainTheirTextStrutAnchor() =>
        DocxInlineCompositionTests.CheckPictureFirstTables(new PinnedResolver());

    public static void PortableRenderingDisplayImageRowsPreserveAdjacentFractionOwnership() =>
        DocxInlineCompositionTests.CheckDisplayImageRows(new PinnedResolver());

    public static void PortableRenderingMixedDisplaySourceRowsPreserveInlineGeometry() =>
        DocxInlineCompositionTests.CheckMixedDisplaySourceRows(new PinnedResolver());

    public static void PortableRenderingMixedDisplayCountersPreserveTotalPageDependencies() =>
        DocxInlineCompositionTests.CheckTotalPageMath(new PinnedResolver(), explicitSource: true);

    public static void PortableRenderingTotalPageMathUsesFinalizedCountAndPreservesSource() =>
        DocxInlineCompositionTests.CheckTotalPageMath(new PinnedResolver());

    public static void PortableRenderingTotalPageMathRetainsUnstableCountFallback() =>
        DocxInlineCompositionTests.CheckUnstableTotalPageMath(new PinnedResolver());

    public static void PortableRenderingTotalPageMathRecomputesNormalNotePlacement()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(Case("CompositionProbes/math-total-pages-simple-body.docx"), mode);
            var notes = DocxTests.ReadDocx(Case("CompositionProbes/math-note-content-" + kind + "-wrapped.docx"), mode);
            var document = source with
            {
                BodyElements = source.BodyElements.Take(2).Concat(notes.BodyElements).Concat(source.BodyElements.Skip(2)).ToArray(),
                RelatedStories = notes.RelatedStories
            };
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            TestAssert.Equal(2, pages.Length);
            string text = Regex.Replace(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)), @"\s", "");
            TestAssert.Contains("Endofsupportednote.", text);
            TestAssert.Equal(4, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            var first = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], usePageCoordinates: true).ToArray();
            TestAssert.Equal(2, first.Count(show => show.Text.Trim() == "2"));
            TestAssert.True(first.Any(show => show.Text == "r"), "The total-page retry must restore the fractions while retaining normal-note contents.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingFooterTableAdmissionKeepsBodyAboveTheCompleteFrame()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            // The pinned face has ordinary math letters, not Cambria's italic
            // mathematical alphabet. Keep this portable geometry control normal.
            var source = ReadPortableNormalMath("Composition/composed-static-cell-report.docx", mode);
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            foreach (var page in renderer.InspectLayout(source).Pages)
            {
                var rows = page.StaticStories.Single(story => story.Kind == "Footer").Items.Where(item => item.Kind.EndsWith("TableRow", StringComparison.Ordinal)).ToArray();
                TestAssert.Equal(1, rows.Length);
                double footerTop = rows.Max(row => row.Y + row.Height);
                var bodyLines = page.Items.Where(item => item.Kind == "TextLine").ToArray();
                TestAssert.True(bodyLines.Length != 0 && bodyLines.All(line => line.Y > footerTop),
                    $"Body admission reserves the complete static table: {mode}, body={bodyLines.Min(line => line.Y)}, footer={footerTop}, margin={page.MarginBottom}.");
            }
            foreach (var page in renderer.InspectTextEmission(source).Lines.GroupBy(line => line.PageIndex))
            {
                var staticSegments = page.Where(line => line.IsStaticStory).SelectMany(line => line.Segments).ToArray();
                TestAssert.True(staticSegments.Length != 0 && staticSegments.All(segment => segment.BaselineY > 0d && segment.BaselineY < source.PageHeightPoints),
                    $"Static content must remain on the printed page: {mode}, page={page.Key}, min={staticSegments.Min(segment => segment.BaselineY)}, max={staticSegments.Max(segment => segment.BaselineY)}.");
                var footer = staticSegments.Where(segment => segment.BaselineY < source.PageHeightPoints / 2d).ToArray();
                var body = page.Where(line => !line.IsStaticStory && line.ContainerStoryKind == "Body").SelectMany(line => line.Segments).ToArray();
                TestAssert.True(footer.Length != 0 && body.Length != 0 &&
                    body.Min(segment => segment.BaselineY) > footer.Max(segment => segment.BaselineY + segment.PdfFontSize),
                    "Printed body baselines clear the visible footer block after both frames are transformed.");
            }
        }
    }

    public static void PortableRenderingStaticCellsPreserveMixedMathImagesAndCounters()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-static-cell-composition.docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string owner in new[] { "Body", "Cell", "Header", "Footer" })
                TestAssert.Contains(owner + ":", text);
            string compact = Regex.Replace(text, @"\s", "");
            TestAssert.Equal(8, compact.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(4, compact.Split("Page1", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(2, compact.Split("Page1of1", StringSplitOptions.None).Length - 1);
            TestAssert.True(pages[0].Images.Count > 0, "Shared images retain their PDF resources.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Moving a supported mixed paragraph into a simple page-owned cell cannot omit its equations.");
            foreach (string part in new[] { "/word/static-cell-header.xml", "/word/static-cell-footer.xml" })
                TestAssert.Equal(1, diagnostics.Count(item => item.Id == "DOCX_APPROXIMATED_EQUATION" && item.PartName == part));
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingStaticCellCountersUseActualPageAndTotalAcrossDigitBoundary()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(Case("CompositionProbes/math-static-cell-composition.docx"), mode);
            var body = new List<DocxBodyElement>(source.BodyElements);
            for (int index = 1; index < 12; index++)
            {
                body.Add(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, null, null));
                body.Add(new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Continuation.", 12d, 12d)));
            }
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(source with { BodyElements = body }, null, default).ToArray();
            TestAssert.Equal(12, pages.Length);
            foreach (int number in new[] { 9, 10, 12 })
            {
                string text = Regex.Replace(DocxPageFootnoteNumberingTests.DecodeText(pages[number - 1])
                    .Normalize(NormalizationForm.FormKC), @"\s", "");
                TestAssert.Equal(2, text.Split("Page" + number + "of12", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(4, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            }
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingAuthoredWhitespaceSurvivesMixedMathContainers()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-authored-inline-whitespace.docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string container in new[] { "Body", "Cell", "Header" })
                TestAssert.Contains("  " + container + " Leading:", text);
            TestAssert.Equal(18, Regex.Replace(text, @"\s", "").Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Whitespace ownership retains every equation.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingBodyPageCountersFollowLineAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(Case("CompositionProbes/math-body-page-counter-boundary.docx"), mode);
            // The pinned portable font has a shorter math strut than Arial.
            // Give this test's spacer a fixed slot that leaves one line on page 9.
            var spacer = source.Paragraphs.Single(paragraph => paragraph.LineSpacingPoints == 600d);
            var taller = spacer with { LineSpacingPoints = 608d, Spacing = spacer.Spacing with { LineValue = "12160" } };
            var document = source with { BodyElements = source.BodyElements.Select(element =>
                element is DocxParagraphElement paragraph && ReferenceEquals(paragraph.Paragraph, spacer)
                    ? new DocxParagraphElement(taller) : element).ToArray() };
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(10, pages.Length);
            foreach (int number in new[] { 9, 10 })
            {
                string text = Regex.Replace(DocxPageFootnoteNumberingTests.DecodeText(pages[number - 1])
                    .Normalize(NormalizationForm.FormKC), @"\s", "");
                TestAssert.Contains(number == 9 ? "Start9" : "middle10End.", text);
                TestAssert.Equal(1, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                TestAssert.True(!text.Contains("{PAGE}", StringComparison.Ordinal), "Body counters bind after line admission.");
            }
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Pagination cannot remove supported mixed math.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingCellPageCountersPreserveMathAcrossDigitBoundary()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-cell-page-counter-boundary.docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(10, pages.Length);
            foreach (int number in new[] { 9, 10 })
            {
                string text = Regex.Replace(DocxPageFootnoteNumberingTests.DecodeText(pages[number - 1])
                    .Normalize(NormalizationForm.FormKC), @"\s", "");
                TestAssert.Contains("Start" + number + "middle" + number + "End.", text);
                TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                TestAssert.True(!text.Contains("{PAGE}", StringComparison.Ordinal), "Cell fields use the actual fragment page.");
            }
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "PAGE cannot hide supported table math.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingNormalNotesPreserveInlineImagesAndMath()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string suffix in new[] { "", "-wrapped" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/" + kind + "-image-math" + suffix + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = Regex.Replace(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText))
                .Normalize(NormalizationForm.FormKC), @"\s", "");
            TestAssert.Equal(4, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Endofsupportednote.", text.Replace("ratetime", "", StringComparison.Ordinal));
            TestAssert.Equal(1, pages.Single().Images.Count);
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Normal-note pictures and math retain their shared frame.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingNormalNotesKeepSectionDesignWidth()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-note-content-" + kind + "-wrapped.docx"), mode);
            var page = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .InspectLayout(document).Pages.Single();
            var note = page.PlacedRelatedStories.Single(story => story.Id == "1" && story.Type is null);
            double width = page.Width - page.MarginLeft - page.MarginRight + page.MarkupMarginReservePoints;
            TestAssert.True(Math.Abs(note.Width - width) < .001d,
                "Normal notes retain the owning section's full design text frame before the shared print transform.");
        }
    }

    public static void PortableRenderingNormalNoteContentPreservesMixedMath()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string flow in new[] { "inline", "wrapped" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-note-content-" + kind + "-" + flow + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = Regex.Replace(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText))
                .Normalize(NormalizationForm.FormKC), @"\s", "");
            TestAssert.Equal(4, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Publiccalculationnote.", text);
            TestAssert.Contains("Endofsupportednote.", text.Replace("ratetime", "", StringComparison.Ordinal));
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Normal-note math uses the same portable rendering resources.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingAutomaticEndnotesPreserveMixedMathAndDestinations()
    {
        foreach (string position in new[] { "before", "after" })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-endnote-" + position + "-" + container + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Public calculation note.", text);
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Endnote composition preserves supported equations.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingEndnoteStartUsesLocalParagraphSpacing()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(Case("CompositionProbes/math-endnote-before-" + container + ".docx"), mode);
            var trailing = source.Paragraphs.Last();
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            double Gap(int after)
            {
                var paragraph = trailing with { SpacingAfterPoints = after,
                    Spacing = trailing.Spacing with { AfterValue = (after * 20).ToString(System.Globalization.CultureInfo.InvariantCulture) } };
                var document = source with { BodyElements = source.BodyElements.Select(element =>
                    element is DocxParagraphElement p && ReferenceEquals(p.Paragraph, trailing)
                        ? new DocxParagraphElement(paragraph) : element).ToArray() };
                var lines = renderer.InspectTextEmission(document).Lines;
                var body = lines.Last(line => line.ContainerStoryKind == "Body");
                var note = lines.Single(line => line.ContainerStoryKind == "Endnote" && line.TextLength > 1);
                return body.Segments[0].BaselineY - note.Segments.MaxBy(segment => segment.TextLength)!.BaselineY;
            }
            double zero = Gap(0), eight = Gap(8), twentyFour = Gap(24);
            TestAssert.True(eight > zero && Math.Abs((twentyFour - zero) - 3d * (eight - zero)) < .001d,
                "Numeric trailing spacing is charged once in the printed frame, independently of earlier tables.");
        }
    }

    public static void PortableRenderingStaticPageFieldsPreserveMathAcrossDigitBoundary()
    {
        foreach (string form in new[] { "simple", "packed" })
        foreach (string story in new[] { "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-page-fields-" + form + "-" + story + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(12, pages.Length);
            for (int index = 0; index < pages.Length; index++)
            {
                string text = System.Text.RegularExpressions.Regex.Replace(DocxPageFootnoteNumberingTests.DecodeText(pages[index])
                    .Normalize(NormalizationForm.FormKC), @"\s", "");
                TestAssert.True(text.Contains("Start" + (index + 1) + "middle12End.", StringComparison.Ordinal),
                    $"{form}/{story}/{mode}/page {index + 1}: {text}");
                TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                TestAssert.True(!text.Contains("{PAGE}", StringComparison.Ordinal) && !text.Contains("{NUMPAGES}", StringComparison.Ordinal),
                    "Source placeholders cannot leak into printed page content.");
            }
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Page-owned fields cannot remove supported static-story math.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingCachedFieldsPreserveMixedMathAcrossContainers()
    {
        foreach (string form in new[] { "simple", "packed" })
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-fields-" + form + "-" + container + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            foreach (string value in new[] { "North", "South" })
                TestAssert.Equal(2, text.Split(value, StringSplitOptions.None).Length - 1);
            foreach (string value in new[] { "Start", "middle", "End.", "Following prose." }) TestAssert.Contains(value, text);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Cached field metadata cannot discard supported math.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingCommentSelectionUsesOccupiedLineFrame()
    {
        foreach (string container in new[] { "body", "cell" })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-comment-prose-" + container + ".docx"), OoxPdfDocxMarkupMode.AllMarkup);
            var renderer = new DocxRenderer(new PinnedResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var layout = renderer.InspectLayout(document).Pages.Single();
            double occupied = layout.Items.SelectMany(item => item.Kind == "TextLine" ? new[] { item }
                : item.TextLines ?? []).Max(line => line.LineHeightPoints ?? 0d);
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var fills = Regex.Matches(page.Content,
                @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
                .Cast<Match>().Where(match => double.Parse(match.Groups["x"].Value,
                    System.Globalization.CultureInfo.InvariantCulture) < 400d).ToArray();
            TestAssert.Equal(1, fills.Length);
            double height = double.Parse(fills[0].Groups["h"].Value, System.Globalization.CultureInfo.InvariantCulture);
            TestAssert.True(occupied > 0d && height >= occupied - .002d && height <= occupied + .2d,
                $"Selection must contain the occupied frame while allowing visible bounds to enlarge it: {container}/{height}/{occupied}.");
            double bottom = double.Parse(fills[0].Groups["y"].Value, System.Globalization.CultureInfo.InvariantCulture);
            double connector = renderer.InspectMarkupBalloons(document).Single().AnchorY;
            TestAssert.True(Math.Abs(connector - bottom) < .002d,
                $"Range selection and connector must share their measured bottom edge: {container}/{connector}/{bottom}.");
            CheckWrittenPages([page]);
        }
    }

    public static void PortableRenderingReviewPaneIgnoresBodyVerticalMargins()
    {
        (double Y, double Height)? expected = null;
        foreach (string margins in new[] { "36-36", "72-72", "144-144", "36-456", "144-36", "36-144" })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/comment-pane-letter-" + margins + ".docx"), OoxPdfDocxMarkupMode.AllMarkup);
            var page = new DocxRenderer(new PinnedResolver(), OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).Single();
            var bounds = DocxCommentsTests.ReviewLaneBounds(page);
            if (expected is null) expected = (bounds.Y, bounds.Height);
            else TestAssert.True(Math.Abs(bounds.Y - expected.Value.Y) < .000001d && Math.Abs(bounds.Height - expected.Value.Height) < .000001d,
                "Changing body margins cannot change the printed review-paper frame.");
            CheckWrittenPages([page]);
        }
    }

    public static void PortableRenderingOrderedMathCellRetainsPrintedBorderFrame()
    {
        foreach (int size in new[] { 11, 12 })
        foreach (string script in new[] { "baseline", "superscript" })
        {
            double? nominalGap = null;
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(Case($"CompositionProbes/math-note-border-{size}-{script}.docx"), mode);
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var renderer = new DocxRenderer(new PinnedResolver(), mode, geometry);
                var table = renderer.InspectLayout(document).Pages.Single().Items.Single(item => item.Kind == "TableRow");
                double gap = table.Y + table.Height - table.TextLines!.Single(line => line.TextLength == 1).Y;
                if (nominalGap is null) nominalGap = gap;
                else
                {
                    double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document, DocxMarkupContext.FromMode(mode, geometry));
                    TestAssert.True(Math.Abs(gap - nominalGap.Value * scale) < .1d,
                        $"The physical math baseline scales once with its table frame: {size}/{script}/gap={gap}/expected={nominalGap.Value * scale}.");
                }
                var pages = renderer.RenderBlankPages(document, null, default).ToArray();
                TestAssert.Equal(1, pages.Length);
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(1, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                TestAssert.Contains("Public calculation note.", text);
                TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
                CheckWrittenPages(pages);
            }
        }
    }

    public static void PortableRenderingScaledFootnoteSharesPrintedBodyCapacity()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
                "CompositionProbes", "scaled-note-" + container + ".docx")), mode);
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            TestAssert.True(pages.Length == 1, $"{container} / {mode}: expected one page, got {pages.Length}.");
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            for (int section = 1; section <= 6; section++)
            for (int line = 1; line <= 3; line++)
                TestAssert.True(text.Contains($"Section {section} line {line}.", StringComparison.Ordinal), "Body admission preserves all source lines.");
            TestAssert.True(text.Contains("Public calculation note.", StringComparison.Ordinal), "Note content remains beside its reference page.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingAtomicFootnoteRetainsLastBodyLine()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/complete-note-" + container + ".docx"), mode);
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            TestAssert.True(text.Contains("Section 6 line 3.", StringComparison.Ordinal) &&
                text.Contains("Public calculation note.", StringComparison.Ordinal), "The final body line and atomic note share the admitted page.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingFootnoteLinksIgnoreIdenticalAuthoredDigits()
    {
        foreach (bool includeMath in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(Case("CompositionProbes/math-footnote-before-body-plain.docx"), mode);
            var paragraph = source.Paragraphs.Single(p => p.Equations.Count != 0);
            var owner = paragraph.InlineReferences.Single().DisplayRun!;
            var authored = owner with { IsAutomaticNoteReferenceMarker = false };
            var runs = paragraph.Runs.ToList();
            runs.Insert(runs.FindIndex(run => ReferenceEquals(run, owner)) + 1, authored);
            var items = paragraph.InlineContent!.ToList();
            items.Insert(items.FindIndex(item => item is DocxInlineText text && ReferenceEquals(text.Run, owner)) + 1,
                new DocxInlineText(authored, 0, authored.Text.Length));
            var changed = paragraph with { Runs = runs, Equations = includeMath ? paragraph.Equations : [],
                InlineContent = includeMath ? items : items.Where(item => item is not DocxInlineEquation).ToArray() };
            var document = source with { BodyElements = source.BodyElements.Select(element =>
                element is DocxParagraphElement p && ReferenceEquals(p.Paragraph, paragraph)
                    ? new DocxParagraphElement(changed) : element).ToArray() };
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
            TestAssert.Equal(includeMath ? 2 : 0, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("ratetime", StringSplitOptions.None).Length - 1);
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingAutomaticFootnotesPreserveMixedMathAndLinks()
    {
        foreach (string position in new[] { "before", "after" })
        foreach (string container in new[] { "body", "cell" })
        foreach (string suffix in new[] { "", "-plain" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-footnote-" + position + "-" + container + suffix + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Public calculation note.", text);
            var link = pages.Single().Annotations.Single(annotation => annotation.Destination is not null);
            TestAssert.True(link.Width > 0d && link.Height > 0d && link.Destination!.Value.PageIndex == 0,
                "Portable marker geometry must retain the source-to-note destination.");
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Automatic footnotes cannot remove the mathematical content.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingCommentsPreserveMixedMathematics()
    {
        foreach (string anchor in new[] { "prose", "between", "formula", "mixed" })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-comment-" + anchor + "-" + container + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            foreach (string word in new[] { "Start", "middle", "End.", "Following prose." }) TestAssert.Contains(word, text);
            if (mode == OoxPdfDocxMarkupMode.AllMarkup) TestAssert.Contains("Check composed calculation.", text);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Source annotation metadata cannot discard the formulas.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingLeftTabsPreserveMixedMathAcrossContainers()
    {
        foreach (string stop in new[] { "default18", "left72" })
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-tabs-" + stop + "-" + container + ".docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            foreach (string field in new[] { "Start", "Middle", "End.", "Following prose." }) TestAssert.Contains(field, text);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "A tab cannot remove supported mathematics in any migrated container.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingDocumentDefaultOwnerPreservesInheritedLineSlots()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var gaps = new Dictionary<string, double>();
            foreach (string scope in new[] { "bare", "paragraph-defaults", "paragraph-defaults-empty", "normal-278-paragraph-defaults" })
            {
                var document = DocxTests.ReadDocx(Case("CompositionProbes/line-spacing-owner-" + scope + ".docx"), mode);
                var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var lines = renderer.InspectLayout(document).Pages.SelectMany(page => page.Items)
                    .Where(item => item.Kind == "TextLine" && item.SourceBlockIndex == 0).ToArray();
                TestAssert.True(lines.Length >= 2, "The source paragraph must wrap to expose its inherited line slot.");
                gaps.Add(scope, lines[0].Y - lines[1].Y);
                var pages = renderer.RenderBlankPages(document, null, default).ToArray();
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
                foreach (int index in Enumerable.Range(1, 6)) TestAssert.Contains("Case " + index + ".", text);
                CheckWrittenPages(pages);
            }
            TestAssert.True(Math.Abs(gaps["paragraph-defaults"] - gaps["paragraph-defaults-empty"]) < 0.000001d,
                "An empty default owner has the same line inheritance as an owner with only side spacing.");
            TestAssert.True(Math.Abs(gaps["bare"] - gaps["normal-278-paragraph-defaults"]) < 0.000001d,
                "An explicit inherited Normal line value must override the document-default fallback.");
            TestAssert.True(gaps["paragraph-defaults"] < gaps["bare"], "Source-owned single spacing must not silently inherit implicit Normal leading.");
        }
    }

    public static void PortableRenderingBrokenLiteralsSurviveCellAndStaticPlacement()
    {
        foreach (string story in new[] { "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("PortableFonts/math-literal-inline-break-default-" + story + "-upright.docx"), mode);
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string glyph in new[] { "α", "+", "β", "=", "γ" })
                TestAssert.Equal(1, text.Split(glyph, StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Following prose.", text);
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingInlineLiteralOperatorBreakPreservesContent()
    {
        foreach (string setting in new[] { "default", "before", "after", "repeat" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("PortableFonts/math-literal-inline-break-" + setting + "-upright.docx"), mode);
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string glyph in new[] { "α", "+", "β", "=", "γ" })
                TestAssert.Equal(1, text.Split(glyph, StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Before", text);
            TestAssert.Contains("Following prose.", text);
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingNumberedMathPreservesFourStoryContent()
    {
        foreach (string story in new[] { "", "-cell", "-header", "-footer" })
        foreach (bool numbered in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("PortableFonts/math-composition-" + (numbered ? "numbered" : "unnumbered") + story + "-normal.docx"), mode);
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
            foreach (string label in new[] { "1.", "2.", "3." })
                TestAssert.Equal(numbered ? 1 : 0, text.Split(label, StringSplitOptions.None).Length - 1);
            TestAssert.Contains("share the same input series.", text);
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingEquationInspectionPreservesGeometryAndSource()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("PortableFonts/math-navigation-normal.docx"), mode);
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var snapshot = renderer.InspectLayout(document);
            var page = snapshot.Pages.Single();
            var equation = page.Items.SingleOrDefault(item => item.Kind == "Equation");
            TestAssert.True(equation is not null, "A painted standalone equation must have an inspectable layout frame.");
            TestAssert.True(equation!.Width > 0d && equation.Height > 0d && equation.LineHeightPoints > 0d &&
                double.IsFinite(equation.X) && double.IsFinite(equation.Y), "Equation geometry must be finite and measurable.");
            TestAssert.True(equation.X >= 0d && equation.X + equation.Width <= page.Width &&
                equation.Y > 0d && equation.Y + equation.Height < page.Height, "The equation frame must lie on its page.");
            TestAssert.True(equation.BaselineYPoints >= equation.Y &&
                equation.BaselineYPoints <= equation.Y + equation.Height, "Inspection must distinguish baseline from box edges.");
            TestAssert.True(equation.SourceBlockIndex is not null && equation.ColumnIndex == 0,
                "Equation inspection must retain source and column ownership.");
            var source = snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == equation.SourceBlockIndex);
            TestAssert.Equal("Paragraph", source.Kind);
            TestAssert.True(source.VerticalTop > source.VerticalBottom && source.ItemCount == 1,
                "The source block must include the equation's nonempty frame.");
            TestAssert.True(page.Items.All(item => item.Kind != "Unknown"), "This supported document must have no unknown layout items.");
        }
    }

    public static void PortableRenderingMathRowKeepsPaintedContentBeforeUnpaintedTail()
    {
        foreach (string suffix in new[] { "0", "6", "24", "24-cant-split", "24-tight" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("PortableFonts/math-table-boundary-after" + suffix + "-normal.docx"), mode);
            var pages = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            int expectedPage = suffix.Contains("cant-split", StringComparison.Ordinal) || suffix.Contains("tight", StringComparison.Ordinal) ? 1 : 0;
            string text = DocxPageFootnoteNumberingTests.DecodeText(pages[expectedPage]).Normalize(NormalizationForm.FormKC);
            TestAssert.Contains("x+2y", text);
            TestAssert.Contains("Observed series", text);
            TestAssert.Equal(1, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("x+2y", StringSplitOptions.None).Length - 1);
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingAutofitKeepsMathematicalCellContent()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var document = DocxTests.ReadDocx(Case("PortableFonts/math-cell-spacing-normal.docx"), mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(4, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION" || d.PartName != document.MainPartName),
                "Autofit must account for prepared math instead of narrowing its column until supported body content disappears.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingAdditionalCellParagraphsPreserveTableOrigin()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var sole = renderer.InspectLayout(DocxTests.ReadDocx(Case("CompositionProbes/math-printed-frame-fraction-anchor12-cell-sole.docx"), mode))
                .Pages.Single().TableRows.Single();
            var flow = renderer.InspectLayout(DocxTests.ReadDocx(Case("CompositionProbes/math-printed-frame-fraction-anchor12-cell-flow.docx"), mode))
                .Pages.Single().TableRows.Single();
            TestAssert.True(flow.Height > sole.Height, "Additional cell paragraphs must contribute occupied height.");
            TestAssert.True(Math.Abs(sole.Y + sole.Height - flow.Y - flow.Height) < .01d,
                "Adding supported cell paragraphs must not shift the top of their table.");
        }
    }

    public static void PortableRenderingIsolatedFractionPreservesCellTextAndFrame()
    {
        foreach (string context in new[] { "sole", "flow" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var document = DocxTests.ReadDocx(Case("CompositionProbes/math-printed-frame-fraction-anchor12-cell-" + context + ".docx"), mode);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(1, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            if (context == "flow")
            {
                TestAssert.Contains("Cell preface.", text);
                TestAssert.Contains("Cell afterword.", text);
            }
            var page = renderer.InspectLayout(document).Pages.Single();
            var cell = page.TableRows.Single().Cells[1];
            var equation = page.Items.Single(item => item.Kind == "TableRow").TextLines!
                .Single(line => line.X > cell.X && line.SourceParagraphIndex == (context == "flow" ? 1 : 0));
            TestAssert.True(equation.Y >= cell.Y && equation.Y <= cell.Y + cell.Height,
                "A placed mathematical baseline must stay inside its admitted cell frame.");
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingFittingMathPreservesOrdinaryLineAdvance()
    {
        string source = Case("CompositionProbes/math-printed-frame-literal-anchor12.docx");
        using var plainBytes = new MemoryStream();
        using (var input = ZipFile.OpenRead(source))
        using (var output = new ZipArchive(plainBytes, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var entry in input.Entries)
            {
                using var data = entry.Open();
                using var target = output.CreateEntry(entry.FullName).Open();
                if (entry.FullName != "word/document.xml") data.CopyTo(target);
                else
                {
                    var xml = XDocument.Load(data);
                    XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    foreach (var equation in xml.Descendants(m + "oMath").ToArray())
                        equation.ReplaceWith(new XElement(w + "r", new XElement(w + "t", "reference")));
                    xml.Save(target);
                }
            }
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            plainBytes.Position = 0;
            var plain = new DocxReader().Read(OoxPackage.Open(plainBytes, default), null, default, markupMode: mode);
            var math = DocxTests.ReadDocx(source, mode);
            double Following(DocxDocument document) => renderer.InspectLayout(document).Pages.Single().Items
                .Single(item => item.Kind == "TextLine" && item.SourceBlockIndex == 2).Y;
            TestAssert.True(Math.Abs(Following(plain) - Following(math)) < .01d,
                "Mathematical ink inside the ordinary line box must not displace the following paragraph.");
        }
    }

    public static void PortableRenderingRepeatedInlineMathPreservesBaselineAndAdvance()
    {
        foreach (string kind in new[] { "literal", "fraction" })
        foreach (int anchor in new[] { 12, 18 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            string name = "CompositionProbes/math-printed-frame-" + kind + "-anchor" + anchor;
            var single = renderer.InspectLayout(DocxTests.ReadDocx(Case(name + ".docx"), mode)).Pages.Single();
            var multiple = renderer.InspectLayout(DocxTests.ReadDocx(Case(name + "-multiple.docx"), mode)).Pages.Single();
            foreach (int block in new[] { 1, 2 })
            {
                var before = single.Items.Single(item => item.Kind == "TextLine" && item.SourceBlockIndex == block);
                var after = multiple.Items.Single(item => item.Kind == "TextLine" && item.SourceBlockIndex == block);
                TestAssert.True(Math.Abs(before.Y - after.Y) < .01d,
                    $"Repeating a fitting formula must preserve its baseline and following advance: {kind}/{anchor}/{mode}/block{block} delta={after.Y - before.Y}.");
            }
        }
    }

    public static void PortableRenderingCommentsPreserveAuthoredWhitespace() =>
        CheckCommentAuthoredWhitespace(new PinnedResolver());

    internal static void CheckCommentAuthoredWhitespace(IFontResolver? resolver)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case("CompositionProbes/comment-authored-whitespace.docx"), mode);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (string expected in new[] { "Reviewed check. Supporting note1", "Earlier annotation. Separate words",
                "Same run. Following text", "Multiple spaces.  Two spaces", "Nonbreaking gap.\u00a0Following word",
                "After break. Continued words" })
                TestAssert.Contains(expected.Replace('\u00a0', ' '), text.Replace('\u00a0', ' '));
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingCommentLabelsPreserveExplicitInitials()
    {
        var document = DocxTests.ReadDocx(Case("CompositionProbes/comment-labels.docx"), OoxPdfDocxMarkupMode.AllMarkup);
        var pages = new DocxRenderer(new PinnedResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).ToArray();
        string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
        foreach (string label in new[] { "[1]", "[2]", "[RV3]", "[4]", "[QA5]" })
            TestAssert.Contains("Commented " + label + ":", text);
        TestAssert.True(!text.Contains("Public reviewer", StringComparison.Ordinal) && !text.Contains("Alternate author", StringComparison.Ordinal),
            "Word-compatible captions consume explicit initials and encounter-order numbers, without substituting author names.");
        CheckWrittenPages(pages);
    }

    public static void PortableRenderingSingleMathAnnotationsPreserveFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var linked = DocxTests.ReadDocx(Case("PortableFonts/letter-normal-math.docx"), mode);
            var source = linked.BodyElements.OfType<DocxParagraphElement>().First(element =>
                element.Paragraph.Equations.Count == 1 && element.Paragraph.Hyperlinks.Count > 0);
            var plain = linked with { BodyElements = linked.BodyElements.Select(element => ReferenceEquals(element, source)
                ? new DocxParagraphElement(source.Paragraph with { Hyperlinks = [] }) : element).ToArray() };
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.InspectLayout(plain).Pages.SelectMany(page => page.Items).ToArray();
            var after = renderer.InspectLayout(linked).Pages.SelectMany(page => page.Items).ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
                TestAssert.True(before[index].X == after[index].X && before[index].Y == after[index].Y &&
                    before[index].Width == after[index].Width && before[index].Height == after[index].Height,
                    "Hyperlink metadata cannot select different paragraph measurement or placement.");
        }
    }

    public static void PortableRenderingHardBreakMathRowsPreserveDisplayFramesAcrossStories()
    {
        string[] stories = ["body", "cell", "header", "footer"];
        for (int story = 0; story < stories.Length; story++)
        {
            string suffix = story == 0 ? "" : "-" + stories[story];
            string fixture = "PortableFonts/inline-math-navigation-normal" + suffix + ".docx";
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                double width = renderer.InspectLayout(DocxTests.ReadDocx(Case("PortableFonts/math-navigation-normal.docx"), mode))
                    .Pages.Single().Items.Single(item => item.Kind == "Equation").Width;
                var page = renderer.InspectLayout(DocxTests.ReadDocx(Case(fixture), mode)).Pages.Single();
                var items = page.Items.Concat(page.Items.SelectMany(item => item.TextLines ?? []))
                    .Concat(page.StaticItems);
                var brokenRow = items.Single(item => item.Kind.EndsWith("TextLine", StringComparison.Ordinal) &&
                    item.TextLength == 0 && item.SourceLineIndex == 1);
                TestAssert.True(Math.Abs(brokenRow.Width - width) < 0.000001d,
                    "A formula isolated by an authored break must retain its display measurement across containers.");
            }
            CheckMathNavigation(fixture, 8, story);
        }
    }

    public static void PortableRenderingMathOnlyGroupsPreserveDisplayMeasurementAndAlignment()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var single = renderer.InspectLayout(DocxTests.ReadDocx(Case("PortableFonts/math-navigation-normal.docx"), mode));
            double width = single.Pages.Single().Items.Single(item => item.Kind == "Equation").Width;
            var document = DocxTests.ReadDocx(Case("PortableFonts/inline-math-navigation-normal.docx"), mode);
            var grouped = renderer.InspectLayout(document);
            double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document, DocxMarkupContext.FromMode(mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
            // The pinned 11pt default contributes three math units between Inner atoms.
            double fractionGap = 11d / 6d * scale;
            var page = grouped.Pages.Single();
            var staticGroups = page.StaticItems.Where(item => item.Kind.EndsWith("TextLine", StringComparison.Ordinal) && item.TextLength == 0).ToArray();
            TestAssert.Equal(2, staticGroups.Length);
            foreach (var group in staticGroups)
            {
                TestAssert.True(Math.Abs(group.Width - 2d * width - fractionGap) < 0.000001d,
                    "Adjacent fractions retain each display measurement and mathematical spacing: " + group.Width + "/" + width + "/" + fractionGap);
                double center = (page.MarginLeft + page.Width - page.MarginRight) / 2d;
                TestAssert.True(Math.Abs(group.X + group.Width / 2d - center) < 0.000001d,
                    "Math-only rows must center the complete measured group in their container.");
            }
        }
    }

    public static void PortableRenderingMathBookmarksPreserveFourStoryDestinations() =>
        CheckMathNavigation("PortableFonts/math-navigation-normal.docx", 4);

    public static void PortableRenderingMixedMathBookmarksPreserveEightStoryDestinations() =>
        CheckMathNavigation("PortableFonts/inline-math-navigation-normal.docx", 8);

    private static void CheckMathNavigation(string fixture, int count, int brokenPair = 0)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case(fixture), mode);
            var renderer = new DocxRenderer(new PinnedResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            var destinations = pages.SelectMany(page => page.Annotations).Where(annotation => annotation.Destination is not null)
                .Select(annotation => annotation.Destination!.Value).ToArray();
            TestAssert.Equal(count, destinations.Length);
            TestAssert.True(destinations.All(destination => destination.PageIndex == 0 && destination.Top > 0d), "All eight targets must resolve on their emitted page.");
            TestAssert.Equal(count, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("x+2y", StringSplitOptions.None).Length - 1);
            if (count == 8)
            {
                for (int index = 0; index < count; index += 2)
                {
                    if (index == 2 * brokenPair)
                        TestAssert.True(destinations[index].Top > destinations[index + 1].Top,
                            "The explicit break must move its second mathematical target down.");
                    else
                        TestAssert.True(destinations[index + 1].Left > destinations[index].Left,
                            "Adjacent formulas must retain separate source targets in each story.");
                }
                TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
            }
            CheckWrittenPages(pages);
        }
    }

    public static void PortableRenderingDocxFinalPreservesMixedContent() =>
        CheckDocx(OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

    public static void PortableRenderingDocxAllMarkupPreservesMixedContent() =>
        CheckDocx(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    public static void PortableRenderingPptxPreservesCompositeChartAndText()
    {
        var package = OoxPackage.Open(Case("pptx-ladder-11-composite-chart-port.pptx"), default);
        var document = new PptxReader().Read(package, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = new PptxRenderer(new PinnedResolver()).RenderPages(document, package, diagnostics.Add, default).ToArray();
        TestAssert.Equal(1, pages.Length);
        string text = DocxPageFootnoteNumberingTests.DecodeText(pages.Single());
        TestAssert.Contains("Sales Overview", text);
        TestAssert.Contains("Category 1", text);
        TestAssert.Contains("Category 4", text);
        TestAssert.Contains("Series 3", text);
        TestAssert.True(pages[0].Content.Contains(" rg", StringComparison.Ordinal), "The composite chart must also paint graphics.");
        CheckWrittenPages(pages);
    }

    private static void CheckDocx(OoxPdfDocxMarkupMode mode, OoxPdfDocxMarkupGeometryMode geometry)
    {
        var document = DocxTests.ReadDocx(Case("PortableFonts/letter-normal-math.docx"), mode);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var pages = new DocxRenderer(new PinnedResolver(), mode, geometry)
            .RenderBlankPages(document, diagnostics.Add, default).ToArray();
        string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("Engineering review letter", text);
        TestAssert.Contains("Acceptance of the revised delivery schedule", text);
        TestAssert.Contains("All quantities in this synthetic example are illustrative.", text);
        TestAssert.Equal(3, text.Split("x+2y", StringSplitOptions.None).Length - 1);
        TestAssert.True(text.Split("a+b=c", StringSplitOptions.None).Length - 1 >= 3,
            "Body, cell and header equations must reach painting with portable fonts.");
        TestAssert.Equal(1, pages.Sum(page => page.Images.Sum(image => Regex.Matches(page.Content,
            "/" + Regex.Escape(image.ResourceName) + @"\s+Do(?:\s|$)").Count)));
        TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Supported equations must not disappear.");
        CheckWrittenPages(pages);
    }

    private static void CheckWrittenPages(PdfPage[] pages)
    {
        TestAssert.True(pages.Length > 0 && pages.All(page => page.Fonts.Count > 0), "Real embedded text resources are required.");
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, pages, default);
        string pdf = Encoding.Latin1.GetString(output.ToArray());
        TestAssert.Contains("/FontFile2", pdf);
        TestAssert.Contains("/ToUnicode", pdf);
        TestAssert.Contains("%%EOF", pdf);
    }

    private static DocxDocument ReadPortableNormalMath(string relative, OoxPdfDocxMarkupMode mode)
    {
        using var bytes = new MemoryStream();
        using (var source = ZipFile.OpenRead(Case(relative)))
        using (var target = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true))
        {
            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            foreach (var entry in source.Entries)
            {
                using var input = entry.Open();
                using var output = target.CreateEntry(entry.FullName).Open();
                if (entry.FullName.StartsWith("word/", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Load(input);
                    foreach (var run in xml.Descendants(math + "r"))
                    {
                        var properties = run.Element(math + "rPr");
                        if (properties is null) { properties = new XElement(math + "rPr"); run.AddFirst(properties); }
                        var normal = properties.Element(math + "nor");
                        if (normal is null) { normal = new XElement(math + "nor"); properties.Add(normal); }
                        normal.SetAttributeValue(math + "val", "1");
                    }
                    xml.Save(output);
                }
                else { input.CopyTo(output); }
            }
        }
        bytes.Position = 0;
        var package = OoxPackage.Open(bytes, default);
        return new DocxReader().Read(package, null, default, markupMode: mode);
    }

    private static string Case(string relative) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", relative));

    private sealed class PinnedResolver : IFontResolver
    {
        private readonly FileFontProgramSource source;
        private readonly string family;

        public PinnedResolver(bool mathFont = false)
        {
            string path = Case(mathFont ? "PortableFonts/DejaVuMathTeXGyre.ttf" : "PortableFonts/DejaVuSans.ttf");
            TestAssert.Equal(mathFont ? "75A6CB527B0175F9D13AF29C33A6915AE38645016C6FE9881653155CD1F62485"
                : "AE7B7855E115A5966D8B1B3F80F254CCC117EC86F9965E202EE2940453837280",
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            source = new(path);
            family = mathFont ? "DejaVu Math TeX Gyre" : "DejaVu Sans";
        }

        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, family,
            new FontStyleKey(HasMathTable: true), source, IsFallback: false);
    }
}
