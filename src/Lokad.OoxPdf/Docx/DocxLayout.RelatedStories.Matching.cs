using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IEnumerable<DocxInlineReferenceLocation> ResolveReferencedRelatedStoryLocations(
        DocxDocument document,
        DocxRelatedStoryKind kind,
        HashSet<(DocxRelatedStoryKind Kind, string Id)> placedStoryKeys)
    {
        for (int sourceBlockIndex = 0; sourceBlockIndex < document.BodyElements.Count; sourceBlockIndex++)
        {
            foreach (DocxInlineReferenceLocation location in EnumerateInlineReferenceLocations(document.BodyElements, sourceBlockIndex))
            {
                DocxInlineReference reference = location.Reference;
                if (reference.Kind != kind ||
                    reference.Id is null ||
                    !placedStoryKeys.Add((reference.Kind, reference.Id)))
                {
                    continue;
                }

                yield return location;
            }
        }
    }

    private static bool TryResolveReferencedRelatedStoryLayout(
        IReadOnlyDictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> storyByKey,
        DocxInlineReferenceLocation location,
        [NotNullWhen(true)] out DocxRelatedStoryLayout? storyLayout)
    {
        storyLayout = null;
        if (location.Reference.Id is null ||
            !storyByKey.TryGetValue((location.Reference.Kind, location.Reference.Id), out DocxRelatedStoryLayout? resolvedLayout) ||
            resolvedLayout is null)
        {
            return false;
        }

        storyLayout = resolvedLayout;
        return true;
    }

    // R12: reference and section-end page searches run against the once-per-group
    // page index instead of re-walking page item trees per location.
    private static int ResolveSectionEndEndnotePageIndex(
        IReadOnlyList<DocxBodyElement> elements,
        RelatedStoryPageIndex referenceIndex,
        IReadOnlyList<DocxLayoutPage> pages,
        DocxInlineReferenceLocation location)
    {
        int referencePageIndex = referenceIndex.FindFirstPageWithReference(location);
        if (referencePageIndex < 0 ||
            !string.Equals(pages[referencePageIndex].PageSettings.EndnoteReferenceSettings.PositionValue, "sectEnd", StringComparison.OrdinalIgnoreCase))
        {
            return -1;
        }

        (int startBlockIndex, int endBlockIndex) = ResolveSectionBlockRange(elements, location.SourceBlockIndex);
        int sectionEndPageIndex = referenceIndex.FindLastPageWithBlockInRange(startBlockIndex, endBlockIndex);
        return sectionEndPageIndex >= 0 ? sectionEndPageIndex : referencePageIndex;
    }

    private static Dictionary<int, List<DocxInlineReferenceLocation>> BuildInlineReferenceLocationsByBlock(
        DocxDocument document,
        CancellationToken cancellationToken)
    {
        // R12: reference locations derive from the immutable document, so memoize them
        // by source block instead of re-walking (table) paragraphs on every fragment page.
        var locationsByBlock = new Dictionary<int, List<DocxInlineReferenceLocation>>();
        for (int sourceBlockIndex = 0; sourceBlockIndex < document.BodyElements.Count; sourceBlockIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            List<DocxInlineReferenceLocation> blockLocations = EnumerateInlineReferenceLocations(document.BodyElements, sourceBlockIndex).ToList();
            if (blockLocations.Count != 0)
            {
                locationsByBlock[sourceBlockIndex] = blockLocations;
            }
        }

        return locationsByBlock;
    }

    private static (int StartBlockIndex, int EndBlockIndex) ResolveSectionBlockRange(IReadOnlyList<DocxBodyElement> elements, int sourceBlockIndex)
    {
        int startBlockIndex = 0;
        for (int index = Math.Min(sourceBlockIndex - 1, elements.Count - 1); index >= 0; index--)
        {
            if (elements[index] is DocxSectionBreakElement)
            {
                startBlockIndex = index + 1;
                break;
            }
        }

        int endBlockIndex = Math.Max(0, elements.Count - 1);
        for (int index = Math.Max(0, sourceBlockIndex); index < elements.Count; index++)
        {
            if (elements[index] is DocxSectionBreakElement)
            {
                endBlockIndex = index;
                break;
            }
        }

        return (startBlockIndex, endBlockIndex);
    }

    private static double ResolvePageBodyWidth(DocxLayoutPage page)
    {
        return Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight);
    }

    private static double ResolvePlacedStoryHeight(DocxRelatedStoryLayout storyLayout, DocxLayoutPage page)
    {
        return Math.Min(Math.Max(0d, storyLayout.ContentHeight), Math.Max(0d, page.Height - page.MarginTop - page.MarginBottom));
    }

    internal static double ResolveDesignBodyBottomForEndnoteStart(
        double scaledBodyBottom,
        double frameTop,
        double? firstBaselineInset,
        double printScale)
    {
        // RV06 anchor probe (edge-endanchor-5, Word 16.0): document-end endnote
        // placement keys off the design body end, but the body lays out scaled
        // (scaled pitches, unscaled first inset). Inverting the used height with
        // the first-line inset recovers the design bottom exactly under that
        // model (probe: 689.77 -> 683.75 at scale 0.75874); scale 1.0 and
        // degenerate frames keep legacy behavior.
        if (firstBaselineInset is not { } inset ||
            Math.Abs(printScale - 1d) < 0.000000001d)
        {
            return scaledBodyBottom;
        }

        double scaledUsed = frameTop - scaledBodyBottom - inset;
        if (!(scaledUsed > 0d))
        {
            return scaledBodyBottom;
        }

        double designBodyBottom = frameTop - inset - scaledUsed / printScale;
        if (!(designBodyBottom <= scaledBodyBottom))
        {
            return scaledBodyBottom;
        }

        return designBodyBottom;
    }

    private static double ResolveEndnoteStartTop(
        DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placedStories,
        double printScale = 1d)
    {
        double bodyBottom = page.Items.Count == 0
            ? page.Height - page.MarginTop
            : page.Items.Min(item => GetVerticalBounds(item).Y);
        double? firstBaselineInset = page.Items.OfType<DocxTextLineLayout>().Select(line => (double?)line.BaselineY).FirstOrDefault() is { } firstBaseline
            ? page.Height - page.MarginTop - firstBaseline
            : null;
        double designBodyBottom = ResolveDesignBodyBottomForEndnoteStart(
            bodyBottom,
            page.Height - page.MarginTop,
            firstBaselineInset,
            printScale);
        double placedStoryBottom = placedStories
            .Select(story => Math.Max(page.MarginBottom, story.TopY - story.Height))
            .DefaultIfEmpty(page.Height - page.MarginTop)
            .Min();
        return Math.Min(designBodyBottom, placedStoryBottom) - FootnoteSeparatorGapPoints;
    }

    private static double ResolveFittingDocumentEndnoteStartTop(
        DocxLayoutPage page, IReadOnlyList<DocxPlacedRelatedStoryLayout> placedStories, double fallbackTop,
        IReadOnlyList<DocxRelatedStoryLayout> endnotes, DocxRelatedStoryLayout? separator,
        IDocxTextMeasurer? measurer, double printScale)
    {
        // Bottom-anchored footnotes are not the end of the body flow. Admit the
        // space above them only when the complete plain endnote block fits.
        // Overflowing, scaled and complex stories retain the continuation path.
        if (Math.Abs(printScale - 1d) > 0.000000001d || endnotes.Count == 0 ||
            endnotes.Any(story => story.TextLines.Count == 0 || story.TableRows.Count != 0 ||
                story.InlineImages.Count != 0 || story.FloatingDrawings.Count != 0)) return fallbackTop;
        DocxPlacedRelatedStoryLayout[] footnotes = placedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote).ToArray();
        if (footnotes.Length == 0) return fallbackTop;
        double candidateTop = ResolveEndnoteStartTop(page, placedStories
            .Where(story => story.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote).ToArray(), printScale);
        if (CanUseNominalDocumentEndnoteFlow(page, separator, printScale, endnotes))
        {
            candidateTop -= ResolveTrailingBodyAfterSpacing(page);
        }
        double requiredHeight = endnotes.Sum(story => Math.Max(0d, story.ContentHeight)) +
            endnotes.Count * FootnoteSeparatorGapPoints;
        if (separator is not null)
        {
            (DocxTextRun? mark, double size) = FindSeparatorMarkFont(separator.TextLines);
            double gap = Math.Max(ResolveSeparatorGapPoints(mark, size, measurer),
                endnotes.Max(story => ResolveFootnoteContentGapPoints(story.TextLines, mark, size, measurer)));
            requiredHeight += ResolveDocumentEndnoteSeparatorHeight(separator, page, measurer, endnotes, printScale) + gap;
        }
        return candidateTop - requiredHeight >= footnotes.Max(story => story.TopY) ? candidateTop : fallbackTop;
    }

    private static bool CanUseNominalDocumentEndnoteFlow(
        DocxLayoutPage page, DocxRelatedStoryLayout? separator, double printScale,
        IReadOnlyList<DocxRelatedStoryLayout> contentStories)
    {
        if (Math.Abs(printScale - 1d) > 0.000000001d || separator is null ||
            separator.TextLines.Count != 1 || !string.IsNullOrWhiteSpace(separator.TextLines[0].Text) || separator.TableRows.Count != 0 ||
            separator.InlineImages.Count != 0 || separator.FloatingDrawings.Count != 0 ||
            separator.Story.BodyElements.Count != 1 ||
            contentStories.Any(story => story.TableRows.Count != 0 || story.InlineImages.Count != 0 || story.FloatingDrawings.Count != 0) ||
            page.ColumnFrames.Count > 1 ||
            page.Items.Count == 0 || page.Items.Any(item => item is not DocxTextLineLayout))
        {
            return false;
        }

        return page.Items.MinBy(item => GetVerticalBounds(item).Y) is DocxTextLineLayout
            { SourceParagraph: { } paragraph } &&
            paragraph.Spacing.AfterLinesValue is null && paragraph.Spacing.AfterAutoSpacingValue is null &&
            double.IsFinite(paragraph.EffectiveProperties.SpacingAfterPoints);
    }

    private static double ResolveTrailingBodyAfterSpacing(DocxLayoutPage page) =>
        page.Items.MinBy(item => GetVerticalBounds(item).Y) is DocxTextLineLayout { SourceParagraph: { } paragraph }
            ? Math.Max(0d, paragraph.EffectiveProperties.SpacingAfterPoints) : 0d;

    private static double ResolveDocumentEndnoteSeparatorHeight(
        DocxRelatedStoryLayout separator, DocxLayoutPage page, IDocxTextMeasurer? measurer,
        IReadOnlyList<DocxRelatedStoryLayout> contentStories, double printScale)
    {
        double height = ResolveSizeDrivenSeparatorHeight(separator, page, measurer, contentStories);
        // The empty separator carries eight latent points in the legacy story
        // layout. Nominal document-end flow uses the last body paragraph's actual
        // after-spacing instead; separator paragraph spacing is ignored by Office.
        return CanUseNominalDocumentEndnoteFlow(page, separator, printScale, contentStories)
            ? Math.Max(0d, height - DocxDefaults.DefaultParagraphAfterSpacingPoints) : height;
    }

    private static IEnumerable<int> EnumeratePageSourceBlockIndexes(DocxLayoutPage page)
    {
        return page.Items
            .Select(GetSourceBlockIndex)
            .Where(index => index is not null)
            .OfType<int>()
            .Distinct()
            .OrderBy(index => index);
    }

    private static IEnumerable<DocxInlineReferenceLocation> EnumerateInlineReferenceLocations(IReadOnlyList<DocxBodyElement> elements, int sourceBlockIndex)
    {
        if (sourceBlockIndex < 0 || sourceBlockIndex >= elements.Count)
        {
            yield break;
        }

        foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateDirectParagraphs([elements[sourceBlockIndex]]))
        {
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                yield return new DocxInlineReferenceLocation(sourceBlockIndex, paragraph, reference);
            }
        }

        if (elements[sourceBlockIndex] is DocxTableElement tableElement)
        {
            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateTableParagraphs(tableElement.Table))
            {
                foreach (DocxInlineReference reference in paragraph.InlineReferences)
                {
                    yield return new DocxInlineReferenceLocation(sourceBlockIndex, paragraph, reference);
                }
            }
        }
    }

    private static IEnumerable<DocxPageTextLineOwner> EnumeratePageTextLineOwners(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            foreach (DocxPageTextLineOwner owner in EnumerateTextLineOwners(item, inheritedSourceBlockIndex: null))
            {
                yield return owner;
            }
        }
    }

    private static IEnumerable<DocxPageTextLineOwner> EnumerateTextLineOwners(DocxLayoutItem item, int? inheritedSourceBlockIndex)
    {
        switch (item)
        {
            case DocxTextLineLayout line:
                yield return new DocxPageTextLineOwner(line, line.SourceBlockIndex ?? inheritedSourceBlockIndex);
                break;
            case DocxTableRowLayout row:
                int? rowSourceBlockIndex = row.Table.SourceBlockIndex;
                foreach (DocxTableCellLayout cell in row.Cells)
                {
                    foreach (DocxTextLineLayout line in cell.TextLines)
                    {
                        yield return new DocxPageTextLineOwner(line, line.SourceBlockIndex ?? rowSourceBlockIndex);
                    }

                    foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
                    {
                        foreach (DocxPageTextLineOwner owner in EnumerateTextLineOwners(nestedRow, rowSourceBlockIndex))
                        {
                            yield return owner;
                        }
                    }
                }

                break;
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumeratePageTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            foreach (DocxTextLineLayout line in EnumerateTextLines(item))
            {
                yield return line;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTextLines(DocxLayoutItem item)
    {
        switch (item)
        {
            case DocxTextLineLayout line:
                yield return line;
                break;
            case DocxTableRowLayout row:
                foreach (DocxTableCellLayout cell in row.Cells)
                {
                    foreach (DocxTextLineLayout line in cell.TextLines)
                    {
                        yield return line;
                    }

                    foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
                    {
                        foreach (DocxTextLineLayout line in EnumerateTextLines(nestedRow))
                        {
                            yield return line;
                        }
                    }
                }

                break;
        }
    }
}
