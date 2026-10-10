using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IReadOnlyDictionary<(DocxRelatedStoryKind Kind, string Id), PdfLinkDestination> CreateNoteDestinations(
        DocxLayout layout, DocxMarkupContext markupContext, CancellationToken cancellationToken)
    {
        var destinations = new Dictionary<(DocxRelatedStoryKind Kind, string Id), PdfLinkDestination>();
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage page = layout.Pages[pageIndex];
            foreach (DocxPlacedRelatedStoryLayout placed in page.PlacedRelatedStories)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxRelatedStory story = placed.StoryLayout.Story;
                if (story.Kind is not (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) ||
                    !story.IsNormalStoryType || story.Id is null ||
                    placed.TextLines.Count == 0 && placed.TableRows.Count == 0)
                {
                    continue;
                }

                bool mapped = FloatingTextBoxEmissionMap.TryCreate(markupContext, page.Height, out var map);
                double scale = mapped ? map.PrintScale : 1d;
                double left = (mapped ? map.MapEmissionX(placed.X) : placed.X) - 3d * scale;
                double top = mapped ? map.MapEmissionY(placed.TopY) : placed.TopY;
                // First placement wins when a note continues onto later pages.
                destinations.TryAdd((story.Kind, story.Id), new PdfLinkDestination(pageIndex, left, top, Zoom: null));
            }
        }
        return destinations;
    }

    private static IEnumerable<PdfLinkAnnotation> CreateNoteReferenceAnnotations(
        DocxLayoutPage page, int pageNumber, int pageCount, DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        IReadOnlyDictionary<(DocxRelatedStoryKind Kind, string Id), PdfLinkDestination> destinations,
        CancellationToken cancellationToken)
    {
        if (destinations.Count == 0) yield break;
        DocxMarkupContext pageContext = WithPageTextEmissionXOffset(markupContext, page);
        double scale = ResolveTextEmissionFontScale(pageContext);
        double yOffset = ResolveTextEmissionBaselineOffset(pageContext);
        double xOffset = ResolveTextEmissionXOffset(pageContext);
        bool wordProfile = UsesWordCompatibleAllMarkupTextProfile(pageContext);
        foreach (DocxTextLineLayout line in EnumerateBodyTextLines(page))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (line.SourceParagraph is not { } paragraph || paragraph.InlineReferences.Count == 0) continue;
            var references = paragraph.InlineReferences.Where(reference =>
                reference.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                reference.Id is not null && (!string.IsNullOrEmpty(reference.DisplayText) ||
                    reference.CustomMarkText is not null && paragraph.Revisions.Count == 0 &&
                    paragraph.RevisionRanges.Count == 0 && paragraph.FieldReferences.Count == 0 && paragraph.CommentRanges.Count == 0)).ToArray();
            if (references.Length == 0) continue;
            IReadOnlyList<DocxTextEmissionSegment> segments = CreateTextEmissionSegments(line, fontResources,
                pageNumber, pageCount, scale, yOffset, xOffset,
                ShouldSuppressWordCompatibleCommentReferenceSpacer(pageContext), wordProfile, cancellationToken);
            foreach (DocxInlineReference reference in references)
            {
                if (!destinations.TryGetValue((reference.Kind, reference.Id!), out PdfLinkDestination destination)) continue;
                foreach (DocxTextEmissionSegment segment in segments)
                {
                    // Match actual source ownership; a custom mark covers its
                    // first character, while an automatic label covers all digits.
                    if (segment.IsTerminalLineSpace || segment.SourceTextRunIndex != reference.SourceRunIndex ||
                        segment.SourceTextOffsetInRun != reference.TextOffsetInRun ||
                        segment.Width <= 0d || segment.Resource is null && segment.FallbackFace is null)
                    {
                        continue;
                    }
                    DocxTextEmissionSegment marker = segment;
                    if (reference.DisplayText is { } automatic)
                    {
                        if (segment.Text != automatic) continue;
                    }
                    else if (reference.CustomMarkText is { } custom && segment.Text.StartsWith(custom, StringComparison.Ordinal) &&
                        fontResources.TextMeasurer is { } measurer)
                    {
                        double markWidth = measurer.MeasureText(segment.StyleRun, custom, segment.FontSize);
                        if (markWidth <= 0d) continue;
                        marker = segment with { Text = custom, Width = markWidth };
                    }
                    else continue;
                    double width = ResolveHyperlinkAnnotationWidth(marker, wordProfile);
                    double inset = line.BodyLineBoxBaselineInsetPoints ??
                        DocxLineMetrics.ResolveTableCellFirstBaselineInset([paragraph], fontResources.TextMeasurer);
                    double top = line.BaselineY - yOffset + inset;
                    double height = line.BodyLineBoxHeightPoints ?? line.LineHeight ?? line.FontSize * 1.2d;
                    if (width > 0d && height > 0d)
                    {
                        // Word's note links cover the paragraph slot, including a superscript mark.
                        double padding = 2.25d * scale;
                        yield return PdfLinkAnnotation.ToDestination(segment.X - padding, top - height,
                            width + 2d * padding, height, destination);
                    }
                }
            }
        }
    }
}
